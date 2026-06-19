"""비동기 파싱 잡 저장/실행 (adaptive_chunk service/jobs.py 이식본).

잡 워커는 JobStore 가 소유한 전용 백그라운드 이벤트루프 스레드에서 돈다.
- POST /parse/jobs/file → job_id 즉시 반환.
- GET /parse/jobs/{id} → status/progress/result/error 폴링.
- DELETE /parse/jobs/{id} → 취소(running 이면 cancel).

무거운 parse_excel_for_rag 는 asyncio.to_thread 로 스레드에서 실행 — 워커 루프 비블로킹.
"""

from __future__ import annotations

import asyncio
import threading
import uuid
from concurrent.futures import Future as CFuture
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Awaitable, Callable, Optional


class JobStatus(str, Enum):
    QUEUED = "queued"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCELLED = "cancelled"


_TERMINAL = {JobStatus.SUCCEEDED, JobStatus.FAILED, JobStatus.CANCELLED}


@dataclass
class Job:
    id: str
    status: JobStatus = JobStatus.QUEUED
    progress: float = 0.0
    result: Optional[dict[str, Any]] = None
    error: Optional[str] = None
    task: Optional["asyncio.Task[None]"] = field(default=None, repr=False)

    def to_public(self) -> dict[str, Any]:
        body: dict[str, Any] = {
            "job_id": self.id,
            "status": self.status.value,
            "progress": round(self.progress, 4),
        }
        if self.status is JobStatus.SUCCEEDED:
            body["result"] = self.result
        if self.status in (JobStatus.FAILED, JobStatus.CANCELLED):
            body["error"] = self.error
        return body


JobWork = Callable[[Callable[[float], None]], Awaitable[dict[str, Any]]]


class JobStore:
    def __init__(self, *, timeout_s: float = 300.0) -> None:
        self._jobs: dict[str, Job] = {}
        self.timeout_s = timeout_s
        self._loop = asyncio.new_event_loop()
        self._thread = threading.Thread(
            target=self._loop.run_forever, name="excel-parser-jobs", daemon=True
        )
        self._thread.start()

    def create(self, work: JobWork) -> Job:
        return self._submit(self._create(work))

    def get(self, job_id: str) -> Optional[Job]:
        return self._submit(self._get(job_id))

    def cancel(self, job_id: str) -> Optional[Job]:
        return self._submit(self._cancel(job_id))

    def shutdown(self) -> None:
        self._loop.call_soon_threadsafe(self._loop.stop)

    async def _create(self, work: JobWork) -> Job:
        job = Job(id=uuid.uuid4().hex)
        self._jobs[job.id] = job
        job.task = self._loop.create_task(self._run(job, work))
        return job

    async def _get(self, job_id: str) -> Optional[Job]:
        return self._jobs.get(job_id)

    async def _cancel(self, job_id: str) -> Optional[Job]:
        job = self._jobs.get(job_id)
        if job is None:
            return None
        if job.status in _TERMINAL:
            return job
        if job.task is not None:
            job.task.cancel()
            try:
                await job.task
            except asyncio.CancelledError:
                pass
        return self._jobs.get(job_id)

    async def _run(self, job: Job, work: JobWork) -> None:
        def report(p: float) -> None:
            job.progress = max(0.0, min(1.0, p))

        job.status = JobStatus.RUNNING
        try:
            result = await asyncio.wait_for(work(report), timeout=self.timeout_s)
        except asyncio.CancelledError:
            job.status = JobStatus.CANCELLED
            job.error = "cancelled by client"
            raise
        except asyncio.TimeoutError:
            job.status = JobStatus.FAILED
            job.error = f"timeout after {self.timeout_s}s"
        except Exception as exc:  # noqa: BLE001
            job.status = JobStatus.FAILED
            job.error = f"{type(exc).__name__}: {exc}"
        else:
            job.status = JobStatus.SUCCEEDED
            job.progress = 1.0
            job.result = result

    def _submit(self, coro: Awaitable[Any]) -> Any:
        fut: CFuture = asyncio.run_coroutine_threadsafe(coro, self._loop)  # type: ignore[arg-type]
        return fut.result()
