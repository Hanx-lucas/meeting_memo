"""Check memo text clues; semantic completeness and layout require separate review."""
import argparse
import json
import re
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

FORBIDDEN = [
    "交流口径称", "交流中称", "交流显示", "会议口径称", "会议中称",
    "专家表示", "专家认为", "会上推测", "会上称", "会上有人",
    "交流表明", "本次会议表明", "参会人员认为", "参会人员援引",
    "根据会议内容", "文中表明", "个人推测",
]
TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9+./-]*|[+-]?\d[\d,]*(?:\.\d+)?%?")
HEADING_RE = re.compile(r"^[一二三四五六七八九十]+、")


def load_text(path):
    doc = Document(path)
    # XML order includes paragraphs in tables and nested tables, without duplicates.
    paragraphs = [Paragraph(p, doc) for p in doc.element.body.iter(qn("w:p"))]
    return doc, paragraphs, "\n".join(p.text for p in paragraphs)


def effective_size(run, paragraph):
    if run.font.size is not None:
        return run.font.size.pt
    for style in (run.style, paragraph.style):
        seen = set()
        while style is not None and style.style_id not in seen:
            seen.add(style.style_id)
            if style.font.size is not None:
                return style.font.size.pt
            style = style.base_style
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--docx", required=True)
    ap.add_argument("--source")
    ap.add_argument("--transcript")
    ap.add_argument("--default-format", action="store_true",
                    help="Check built-in heading numbering and 11/10/9 pt sizes only.")
    ap.add_argument("--required-term", action="append", default=[])
    ap.add_argument("--allow-missing-token", action="append", default=[],
                    help="Exact source token intentionally changed; record its reason separately.")
    args = ap.parse_args()
    path = Path(args.docx)
    issues, warnings = [], []
    try:
        with zipfile.ZipFile(path) as zf:
            bad_part = zf.testzip()
            if bad_part:
                raise ValueError(f"Corrupt ZIP part: {bad_part}")
        doc, paragraphs, text = load_text(path)
        source_text = load_text(Path(args.source))[2] if args.source else None
        transcript = (Path(args.transcript).read_text(encoding="utf-8-sig")
                      if args.transcript else None)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
        print(json.dumps({"file": str(path.resolve()), "pass": False,
                          "issues": [str(exc)]}, ensure_ascii=False, indent=2))
        return 1

    if not text.strip():
        issues.append("Document body is empty")
    forbidden_hits = {p: text.count(p) for p in FORBIDDEN if p in text}
    if forbidden_hits:
        issues.append("Review forbidden narrator phrases in context")
    headings = [p.text for p in paragraphs if HEADING_RE.match(p.text)]
    if args.default_format and not headings:
        issues.append("No Chinese level-1 headings found for default format")

    missing_tokens = []
    allowed = set(args.allow_missing_token)
    if source_text is not None:
        missing_tokens = sorted(set(TOKEN_RE.findall(source_text))
                                - set(TOKEN_RE.findall(text)) - allowed)
        if missing_tokens:
            issues.append("Numeric/Latin tokens disappeared from source; review required")
    required_missing = [term for term in args.required_term if term not in text]
    if required_missing:
        issues.append("Required terms are missing")
    transcript_terms_missing = []
    if transcript is not None:
        candidates = set(re.findall(r"[A-Za-z][A-Za-z0-9+./-]{2,}", transcript))
        important = {x for x in candidates if any(c.isupper() for c in x) or len(x) >= 5}
        transcript_terms_missing = sorted(x for x in important if x not in text)
        if transcript_terms_missing:
            warnings.append("Transcript term candidates need contextual review")

    font_findings = []
    nonempty = [p for p in paragraphs if p.text.strip()]
    if args.default_format:
        for i, p in enumerate(nonempty):
            expected = 11 if i == 0 else (10 if HEADING_RE.match(p.text) else 9)
            for r in p.runs:
                if not r.text.strip():
                    continue
                size = effective_size(r, p)
                if size is None:
                    warnings.append(f"Cannot resolve inherited font size in paragraph {i}")
                elif abs(size - expected) > 0.1:
                    font_findings.append({"paragraph": i, "text": r.text[:60],
                                          "size": size, "expected": expected})
        if font_findings:
            issues.append("Font sizes differ from built-in default format")

    report = {
        "file": str(path.resolve()), "paragraphs_including_tables": len(paragraphs),
        "characters": len(text), "headings": headings, "default_format": args.default_format,
        "forbidden_hits": forbidden_hits, "missing_source_tokens": missing_tokens,
        "allowed_missing_tokens": sorted(allowed), "missing_required_terms": required_missing,
        "transcript_terms_not_in_docx_sample": transcript_terms_missing[:100],
        "font_findings": font_findings[:50], "issues": issues,
        "warnings": sorted(set(warnings)), "pass": not issues,
        "scope": "Enabled machine checks only; manually verify content, order, facts and rendered pages.",
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
