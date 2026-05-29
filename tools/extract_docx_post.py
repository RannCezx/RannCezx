from __future__ import annotations

import html
import shutil
import zipfile
from pathlib import Path

from bs4 import BeautifulSoup
from docx import Document
from docx.oxml.ns import qn


def norm(text: str) -> str:
    return " ".join((text or "").replace("\xa0", " ").split()).strip()


def extract_media(docx_path: Path, out_dir: Path) -> None:
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(docx_path) as zf:
        for name in zf.namelist():
            if not name.startswith("word/media/"):
                continue
            target = out_dir / Path(name).name
            with zf.open(name) as src, target.open("wb") as dst:
                shutil.copyfileobj(src, dst)


def build_relationships(doc: Document) -> dict[str, str]:
    rels: dict[str, str] = {}
    for rel_id, rel in doc.part.rels.items():
        target_ref = getattr(rel, "target_ref", "")
        if isinstance(target_ref, str) and "media/" in target_ref:
            rels[rel_id] = Path(target_ref).name
    return rels


def paragraph_image_filenames(para, rels: dict[str, str]) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()
    for node in para._element.iter():
        tag = str(node.tag)
        rid = ""
        if tag.endswith("}blip"):
            rid = node.get(qn("r:embed")) or ""
        elif tag.endswith("}imagedata"):
            rid = node.get(qn("r:id")) or ""
        if not rid:
            continue
        filename = rels.get(rid or "")
        if filename and filename not in seen:
            seen.add(filename)
            names.append(filename)
    return names


def render_html(docx_path: Path, media_web_prefix: str) -> str:
    doc = Document(docx_path)
    rels = build_relationships(doc)
    parts: list[str] = []
    skip_titles = {norm("2026pgs初赛 Write Up"), norm("物联网取证"), norm("物联网取证’"), norm("物联网取证'")}

    for para in doc.paragraphs:
        text = (para.text or "").replace("\xa0", " ").strip()
        clean = norm(text)
        if clean and clean not in skip_titles:
            parts.append(f"<p>{html.escape(text)}</p>")

        for filename in paragraph_image_filenames(para, rels):
            src = f"{media_web_prefix}/{filename}"
            alt = Path(filename).stem
            parts.append(f'<p class="doc-image"><img src="{html.escape(src)}" alt="{html.escape(alt)}"></p>')

    html_text = "\n".join(parts)
    soup = BeautifulSoup(html_text, "html.parser")
    return str(soup)


def main() -> None:
    project_root = Path(__file__).resolve().parents[1]
    docx_path = Path(r"c:\Users\jzz11\Desktop\2026pgs初赛 Write Up.docx")
    assets_dir = project_root / "assets" / "post-1-media"
    html_path = project_root / "assets" / "post-1.html"

    extract_media(docx_path, assets_dir)
    html_text = render_html(docx_path, "assets/post-1-media")
    html_path.write_text(html_text, encoding="utf-8")
    print(f"wrote_html={html_path}")
    print(f"media_dir={assets_dir}")


if __name__ == "__main__":
    main()
