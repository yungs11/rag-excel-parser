"""Form 문서 파싱 검증 — key-value 추출, form_field / form_summary (SoT §15)."""

from __future__ import annotations

import fixture_builders as fb
from conftest import chunk_blob

FORM = "06_form_document.xlsx"
# 사유(긴 문장) 제외, 핵심 4쌍은 반드시 추출되어야 함
CORE_FIELDS = {k: v for k, v in fb.FORM_06_FIELDS.items() if k != "사유"}


class TestFormFields:
    def test_form_field_chunks_generated(self, parse_chunks):
        chunks = parse_chunks(FORM)
        form_fields = [c for c in chunks if c["chunk_type"] == "form_field"]
        assert len(form_fields) >= len(CORE_FIELDS), (
            f"form_field {len(CORE_FIELDS)}개 이상 기대, 실제 {len(form_fields)}개. "
            f"생성된 타입: {sorted({c['chunk_type'] for c in chunks})}"
        )

    def test_each_key_value_extracted(self, parse_chunks):
        """key 와 value 가 같은 form_field chunk 에 함께 있어야 함."""
        chunks = parse_chunks(FORM)
        form_fields = [c for c in chunks if c["chunk_type"] == "form_field"]
        for key, value in CORE_FIELDS.items():
            matched = [
                c for c in form_fields
                if key in chunk_blob(c) and value in chunk_blob(c)
            ]
            assert matched, (
                f"form key-value 미추출: {key}={value}. "
                f"실제 form_field: {[c['content_text'] for c in form_fields]}"
            )

    def test_value_in_content_text(self, parse_chunks):
        """SoT §19.2 — form_field content_text 에 값이 자연어로 포함."""
        chunks = parse_chunks(FORM)
        form_fields = [c for c in chunks if c["chunk_type"] == "form_field"]
        assert any("홍길동" in c["content_text"] for c in form_fields), (
            f"신청자 값이 content_text 에 없음: {[c['content_text'] for c in form_fields]}"
        )

    def test_form_summary_exists(self, parse_chunks):
        """SoT §15.4 — form 전체 요약 chunk 생성."""
        chunks = parse_chunks(FORM)
        assert any(c["chunk_type"] == "form_summary" for c in chunks), (
            f"form_summary 미생성. 생성된 타입: {sorted({c['chunk_type'] for c in chunks})}"
        )

    def test_form_chunk_sources(self, parse_chunks):
        chunks = parse_chunks(FORM)
        for c in chunks:
            if c["chunk_type"] not in ("form_field", "form_summary"):
                continue
            assert c["sheet"] == fb.FORM_06_SHEET
            assert c["source"].get("range"), f"form chunk source.range 누락: {c['id']}"
