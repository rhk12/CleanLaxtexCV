from __future__ import annotations

import argparse
import importlib
import inspect
import re
import subprocess
from shutil import which
from datetime import datetime
from pathlib import Path

from docx import Document

from core.io_utils import ensure_docx, find_latest_dossier, paragraphs_text

FULL_SECTIONS = [
    "sections.professional_positions",
    "sections.education",
    "sections.awards_honors",
    "sections.publications",
    "sections.grants",
    "sections.intellectual_property",
    "sections.directed_student_learning",
    "sections.service",
    "sections.editorial_board",
    "sections.professional_memberships",
]

THREE_PAGE_SECTIONS = [
    "sections.professional_positions",
    "sections.education",
    "sections.awards_honors",
    "sections.publications",
]

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dossier", type=str, default=None)
    parser.add_argument("--template", type=str, default="template_cv.tex")
    parser.add_argument("--output", type=str, default="output.tex")
    parser.add_argument("--header-date", type=str, default=datetime.now().strftime("%B %Y"))
    parser.add_argument("--full-cv", action="store_true")
    parser.add_argument("--three-page-cv", action="store_true")
    parser.add_argument(
        "--compile-pdf",
        action="store_true",
        help="Compile the generated .tex locally into a PDF using latexmk or pdflatex.",
    )
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args()


def resolve_mode(args) -> str:
    if args.three_page_cv:
        return "three-page"
    return "full"


def get_section_modules(mode: str) -> list[str]:
    if mode == "three-page":
        return THREE_PAGE_SECTIONS
    return FULL_SECTIONS


def load_template(path: str, header_date: str) -> str:
    template_path = Path(path)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    text = template_path.read_text(encoding="utf-8")
    return text.replace("{HEADER_DATE}", header_date)


def call_section_apply(module, text: str, doc, document_text: str, mode: str) -> str:
    apply_fn = module.apply
    sig = inspect.signature(apply_fn)
    param_names = list(sig.parameters.keys())

    if len(param_names) == 2:
        return apply_fn(text, doc)

    if len(param_names) == 3:
        return apply_fn(text, doc, document_text)

    if len(param_names) >= 4:
        return apply_fn(text, doc, document_text, mode)

    raise TypeError(f"Unsupported apply() signature in module {module.__name__}: {sig}")


def apply_sections(text: str, doc, document_text: str, mode: str, verbose: bool) -> str:
    section_modules = get_section_modules(mode)

    for module_name in section_modules:
        module = importlib.import_module(module_name)

        if not hasattr(module, "apply"):
            raise AttributeError(f"{module_name} is missing apply(...)")

        if verbose:
            print(f"[info] Applying {module_name}")

        text = call_section_apply(module, text, doc, document_text, mode)

    return text


def strip_unreplaced_placeholders(text: str, verbose: bool = False) -> str:
    placeholders = re.findall(r"\{\{[A-Za-z0-9_]+\}\}", text)

    if verbose and placeholders:
        print("[warn] Unreplaced placeholders found and removed:")
        for ph in placeholders:
            print(f"       {ph}")

    return re.sub(r"\{\{[A-Za-z0-9_]+\}\}\n*", "", text)


def compile_latex(tex_path: Path, verbose: bool = False) -> Path:
    if not tex_path.exists():
        raise FileNotFoundError(f"LaTeX source not found: {tex_path}")

    workdir = tex_path.parent
    build_dir = workdir / "build"
    tex_name = tex_path.name
    build_dir.mkdir(parents=True, exist_ok=True)

    latexmk = which("latexmk")
    pdflatex = which("pdflatex")

    def run_pdflatex_passes() -> None:
        if not pdflatex:
            raise RuntimeError(
                "No LaTeX compiler found. Install latexmk or pdflatex to enable local PDF builds."
            )

        cmd = [
            pdflatex,
            "-interaction=nonstopmode",
            "-halt-on-error",
            "-file-line-error",
            f"-output-directory={build_dir.name}",
            tex_name,
        ]
        if verbose:
            print(f"[info] Compiling with pdflatex: {' '.join(cmd)}")
        subprocess.run(cmd, cwd=workdir, check=True)
        subprocess.run(cmd, cwd=workdir, check=True)

    if latexmk:
        cmd = [
            latexmk,
            "-pdf",
            "-interaction=nonstopmode",
            "-halt-on-error",
            "-file-line-error",
            f"-outdir={build_dir.name}",
            tex_name,
        ]
        if verbose:
            print(f"[info] Compiling with latexmk: {' '.join(cmd)}")
        try:
            subprocess.run(cmd, cwd=workdir, check=True)
        except subprocess.CalledProcessError as exc:
            if verbose:
                print(
                    "[warn] latexmk failed, falling back to pdflatex. "
                    "This usually means MiKTeX needs Perl for latexmk."
                )
            run_pdflatex_passes()
    else:
        run_pdflatex_passes()

    return build_dir / tex_path.with_suffix(".pdf").name


def main():
    args = parse_args()
    mode = resolve_mode(args)

    dossier_path = Path(args.dossier) if args.dossier else find_latest_dossier(Path("Dossiers"))
    dossier_path = ensure_docx(dossier_path)

    if args.verbose:
        print(f"[info] Build mode: {mode}")
        print(f"[info] Output path: {args.output}")
        print(f"[info] Template path: {Path(args.template).resolve()}")
        print(f"[info] Dossier path: {dossier_path}")

    template = load_template(args.template, args.header_date)
    document_text = paragraphs_text(str(dossier_path))
    doc = Document(str(dossier_path))

    output_text = apply_sections(template, doc, document_text, mode, args.verbose)
    output_text = strip_unreplaced_placeholders(output_text, args.verbose)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(output_text, encoding="utf-8")

    if args.verbose:
        print(f"[info] Wrote {output_path} successfully.")

    if args.compile_pdf:
        pdf_path = compile_latex(output_path, args.verbose)
        if args.verbose:
            print(f"[info] Wrote {pdf_path} successfully.")


if __name__ == "__main__":
    main()
