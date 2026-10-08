from pathlib import Path
import ast

ROOT = Path(__file__).resolve().parents[1]


def function_source(path: str, name: str) -> str:
    text = (ROOT / path).read_text(encoding="utf-8")
    tree = ast.parse(text)
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == name:
            return ast.get_source_segment(text, node) or ""
    raise AssertionError(f"Function {name} not found in {path}")


def test_evolution_pdf_hard_blocks_identity_fields():
    obj = function_source("core/pdf_reports.py", "_object_text")
    params = function_source("core/pdf_reports.py", "_parameters_table")
    ident = function_source("core/pdf_reports.py", "_identification_tables")
    assert "participant_name" not in obj
    assert "registration" not in obj
    assert "cpf" not in obj
    assert "Elaboração" not in params
    assert "Validação" not in params
    assert "return None" in ident
    assert "em nome de" not in (ROOT / "core/pdf_reports.py").read_text(encoding="utf-8")


def test_majs_pdf_hard_blocks_participant_and_responsible():
    info = function_source("core/majs_pdf.py", "_info_table")
    footer = function_source("core/majs_pdf.py", "_draw_header_footer")
    assert "participant_name" not in info
    assert "elaborator" not in footer
    assert "validator" not in footer


def test_monetary_pdf_hard_blocks_participant():
    meta = function_source("core/monetary_pdf.py", "_meta_strip")
    assert "participant_name.strip" not in meta
    assert '"Participante"' not in meta
