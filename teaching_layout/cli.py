from __future__ import annotations

import argparse
import json
from pathlib import Path

from .docx_parser import summarize_docx
from .legacy_adapter import build_with_legacy_generator
from .template_inspector import inspect_template
from .template_importer import import_template
from .validator import validate_pdf


def build_command(args: argparse.Namespace) -> None:
    manifest = build_with_legacy_generator(
        template_ref=args.template,
        docx=Path(args.docx),
        output_dir=Path(args.out),
        pdf_name=args.pdf_name,
        preview_name=args.preview_name,
        title=args.title,
        background_mode=args.background_mode,
        templates_dir=Path(args.templates_dir),
    )
    print(json.dumps(manifest["outputs"], ensure_ascii=False, indent=2))


def parse_docx_command(args: argparse.Namespace) -> None:
    summary = summarize_docx(Path(args.docx))
    print(json.dumps(summary, ensure_ascii=False, indent=2))


def validate_command(args: argparse.Namespace) -> None:
    result = validate_pdf(Path(args.pdf), Path(args.manifest) if args.manifest else None)
    print(json.dumps(result, ensure_ascii=False, indent=2))


def import_template_command(args: argparse.Namespace) -> None:
    result = import_template(
        template_id=args.id,
        name=args.name,
        source_dir=Path(args.source),
        templates_dir=Path(args.templates_dir),
        overwrite=args.overwrite,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))


def inspect_template_command(args: argparse.Namespace) -> None:
    result = inspect_template(args.template, templates_dir=Path(args.templates_dir))
    print(json.dumps(result, ensure_ascii=False, indent=2))


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="teaching-layout", description="Local CLI engine for teaching-aid PDF layout.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    build = subparsers.add_parser("build", help="Generate a print PDF from a template package and DOCX.")
    build.add_argument("--template", required=True, help="Template name or package/material root path, e.g. gonggu-neiye")
    build.add_argument("--templates-dir", default=str(Path(__file__).resolve().parents[1] / "templates"))
    build.add_argument("--docx", required=True, help="Source DOCX path")
    build.add_argument("--out", required=True, help="Output directory")
    build.add_argument("--pdf-name", default="output.pdf", help="Output PDF file name")
    build.add_argument("--preview-name", default="preview.png", help="Output preview PNG file name")
    build.add_argument("--title", default="teaching-layout-output", help="PDF metadata title")
    build.add_argument("--background-mode", choices=["white", "image"], default="white")
    build.set_defaults(func=build_command)

    parse_docx = subparsers.add_parser("parse-docx", help="Summarize structured DOCX content.")
    parse_docx.add_argument("--docx", required=True)
    parse_docx.set_defaults(func=parse_docx_command)

    validate = subparsers.add_parser("validate", help="Validate generated PDF and optional manifest.")
    validate.add_argument("--pdf", required=True)
    validate.add_argument("--manifest", default=None)
    validate.set_defaults(func=validate_command)

    import_template_parser = subparsers.add_parser(
        "import-template",
        help="Create a template package skeleton from raw assets.",
    )
    import_template_parser.add_argument("--id", required=True, help="Template id, e.g. gonggu-neiye")
    import_template_parser.add_argument("--name", required=True, help="Human-readable template name")
    import_template_parser.add_argument("--source", required=True, help="Source folder containing assets/")
    import_template_parser.add_argument(
        "--templates-dir",
        default=str(Path(__file__).resolve().parents[1] / "templates"),
        help="Destination templates directory",
    )
    import_template_parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite an existing template package with the same id",
    )
    import_template_parser.set_defaults(func=import_template_command)

    inspect_template_parser = subparsers.add_parser(
        "inspect-template",
        help="Inspect a template package for missing files and local absolute paths.",
    )
    inspect_template_parser.add_argument("--template", required=True, help="Template name, id, or package path")
    inspect_template_parser.add_argument(
        "--templates-dir",
        default=str(Path(__file__).resolve().parents[1] / "templates"),
        help="Templates directory",
    )
    inspect_template_parser.set_defaults(func=inspect_template_command)

    return parser


def main(argv: list[str] | None = None) -> None:
    parser = create_parser()
    args = parser.parse_args(argv)
    args.func(args)
