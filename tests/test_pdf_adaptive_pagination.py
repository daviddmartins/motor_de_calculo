from pathlib import Path


def test_parecer_usa_paginacao_adaptativa_sem_quebras_fixadas():
    source = Path('core/pdf_reports.py').read_text(encoding='utf-8')

    assert 'CondPageBreak' in source
    assert 'story.append(PageBreak())' not in source
    assert "if idx==0: out.append(PageBreak())" not in source
    assert "elif idx>0: out.append(PageBreak())" not in source
    assert 'splitByRow=1' in source


def test_parecer_protege_blocos_essenciais_sem_travar_secoes_inteiras():
    source = Path('core/pdf_reports.py').read_text(encoding='utf-8')

    assert 'conclusion_parts = _conclusion' in source
    assert 'story.append(KeepTogether([' in source
    assert 'out.append(CondPageBreak(72 * mm))' in source
    assert "Paragraph(f'Resumo da Prestação {row.installment_number}'" in source
