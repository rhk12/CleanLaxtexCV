# CleanLaTeXCV Refactor (exact-template version)

This package keeps the CV style locked to `template_cv.tex` and only regenerates section content from the latest Penn State dossier.

## Install

```bash
python -m pip install -r requirements.txt
```

## Full CV

```bash
python main.py --full-cv --verbose
```

## Compact CV

```bash
python main.py --three-page-cv --verbose
```

## Specific dossier

```bash
python main.py --dossier "C:\Users\reube\source\repos\rhk12\CleanLaxtexCV\Dossiers\20260407-083623-CDT.doc.docx" --full-cv --verbose
```

## Force a specific header date

```bash
python main.py --dossier "C:\path\to\file.docx" --full-cv --header-date "December 2024" --verbose
```

## Notes

- The exact visual style comes from `template_cv.tex`.
- To change lines, spacing, rule thickness, preamble packages, or header formatting, edit `template_cv.tex` instead of the Python section files.
- `.doc` files require LibreOffice (`soffice`) for automatic conversion. `.docx` is preferred.
