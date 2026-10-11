"""Render DOCX to PDF and page PNGs using configurable LibreOffice and Poppler."""
import argparse
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def resolve_tool(explicit, env_name, names, windows_paths=()):
    configured = explicit or os.environ.get(env_name)
    if configured:
        candidate = Path(configured).expanduser()
        if candidate.is_file():
            return str(candidate.resolve())
        resolved = shutil.which(configured)
        if resolved:
            return resolved
        raise ValueError(f"{env_name} tool not found: {configured}")
    for name in names:
        resolved = shutil.which(name)
        if resolved:
            return resolved
    if os.name == "nt":
        for variable in ("ProgramFiles", "ProgramFiles(x86)"):
            root = os.environ.get(variable)
            if root:
                for relative in windows_paths:
                    candidate = Path(root) / relative
                    if candidate.is_file():
                        return str(candidate)
    raise ValueError(f"Tool not found: {names[0]}; set {env_name} or its CLI option")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--docx", required=True)
    ap.add_argument("--output-dir", required=True)
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--soffice")
    ap.add_argument("--pdftoppm")
    ap.add_argument("--temp-dir", help="Writable temporary directory for restricted environments.")
    args = ap.parse_args()
    if args.dpi <= 0:
        ap.error("--dpi must be positive")
    source = Path(args.docx).resolve()
    if not source.is_file():
        ap.error(f"DOCX not found: {source}")
    try:
        soffice = resolve_tool(args.soffice, "SOFFICE_PATH", ("soffice.com", "soffice", "soffice.exe"),
                               ("LibreOffice/program/soffice.com", "LibreOffice/program/soffice.exe"))
        pdftoppm = resolve_tool(args.pdftoppm, "PDFTOPPM_PATH", ("pdftoppm", "pdftoppm.exe"))
    except ValueError as exc:
        ap.error(str(exc))
    out = Path(args.output_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)
    run_dir = Path(tempfile.mkdtemp(prefix="render-", dir=out))
    with tempfile.TemporaryDirectory(prefix="memo-render-", dir=args.temp_dir) as temp:
        work = Path(temp)
        ascii_docx = work / "input.docx"
        shutil.copy2(source, ascii_docx)
        profile = work / "lo-profile"
        profile.mkdir()
        command = [soffice, "--headless", f"-env:UserInstallation={profile.as_uri()}",
                   "--convert-to", "pdf", "--outdir", str(work), str(ascii_docx)]
        try:
            result = subprocess.run(command, capture_output=True, text=True, timeout=60)
            pdf = work / "input.pdf"
            if not pdf.exists() or pdf.stat().st_size == 0:
                raise RuntimeError(f"LibreOffice failed ({result.returncode})\n"
                                   f"{result.stdout}\n{result.stderr}")
            final_pdf = run_dir / "input.pdf"
            shutil.copy2(pdf, final_pdf)
            subprocess.run([pdftoppm, "-png", "-r", str(args.dpi), str(final_pdf),
                            str(run_dir / "page")], check=True, timeout=60)
        except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
            raise SystemExit(f"Rendering failed: {exc}; inspect partial output: {run_dir}")
    pages = sorted(run_dir.glob("page-*.png"))
    if not pages:
        raise SystemExit(f"No page PNGs produced: {run_dir}")
    print(f"OUTPUT_DIR={run_dir}")
    print(f"PDF={final_pdf}")
    print(f"PAGES={len(pages)}")


if __name__ == "__main__":
    main()
