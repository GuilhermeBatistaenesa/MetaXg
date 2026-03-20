import os
import re
from datetime import datetime
from PIL import Image as PILImage

from custom_logger import logger
from output_manager import KIND_LOGS, KIND_RELATORIOS, OutputManager
from outcomes import (
    OUTCOME_FAILED_ACTION,
    OUTCOME_FAILED_VERIFICATION,
    OUTCOME_SAVED_NOT_VERIFIED,
    OUTCOME_VERIFIED_SUCCESS,
    OUTCOME_SKIPPED_ALREADY_EXISTS,
    OUTCOME_SKIPPED_DRY_RUN,
    OUTCOME_SKIPPED_NO_RECIPIENT,
    OUTCOME_ORDER,
)


MANUAL_ERROR_OUTCOMES = {
    OUTCOME_FAILED_ACTION,
    OUTCOME_FAILED_VERIFICATION,
    OUTCOME_SAVED_NOT_VERIFIED,
}


def _slug(texto: str) -> str:
    texto = re.sub(r"[^\w\-]+", "_", (texto or "").strip().lower(), flags=re.UNICODE)
    return texto.strip("_") or "sem_texto"


def _nome_base_execucao(run_context: dict) -> str:
    started_at = datetime.fromisoformat(run_context["started_at"])
    status = _slug(run_context.get("run_status") or "running")
    return f"{started_at.strftime('%Y-%m-%d')}__{started_at.strftime('%Hh%M')}__{status}"


def gerar_relatorio_txt(manifest: dict, output_manager: OutputManager):
    """
    Gera um relatorio TXT detalhado da execucao a partir do manifest.
    """
    try:
        run_context = manifest["run_context"]
        started_at = datetime.fromisoformat(run_context["started_at"])
        finished_at = run_context.get("finished_at")
        finished_dt = datetime.fromisoformat(finished_at) if finished_at else None
        duration_sec = run_context.get("duration_sec")
        nome_arquivo = f"{_nome_base_execucao(run_context)}__relatorio_execucao.txt"

        totals = manifest["totals"]
        pessoas = manifest["people"]
        report_path = output_manager.get_local_path(KIND_RELATORIOS, nome_arquivo)

        linhas = []
        linhas.append("RELATORIO DE EXECUCAO RPA METAX")
        linhas.append(f"Execution ID: {run_context['execution_id']}")
        linhas.append(f"Run Status: {run_context['run_status']}")
        linhas.append(f"Started At: {started_at.strftime('%d/%m/%Y %H:%M:%S')}")
        if finished_dt:
            linhas.append(f"Finished At: {finished_dt.strftime('%d/%m/%Y %H:%M:%S')}")
        if duration_sec is not None:
            linhas.append(f"Duration (sec): {duration_sec}")
        linhas.append("=" * 50)
        linhas.append("")
        linhas.append("RESUMO GERAL:")
        linhas.append(f"- Total Detectado no RM: {totals['detected']}")
        linhas.append(f"- Total Processado (people): {totals['people_total']}")
        linhas.append(f"- Sem Foto: {totals['no_photo']}")
        linhas.append("")
        linhas.append("TOTAIS POR OUTCOME:")
        for outcome in OUTCOME_ORDER:
            linhas.append(f"- {outcome}: {totals['by_outcome'].get(outcome, 0)}")
        if totals.get("unknown_outcome"):
            linhas.append(f"- UNKNOWN_OUTCOME: {totals['unknown_outcome']}")
        linhas.append("")

        if run_context.get("public_write_ok") is False:
            linhas.append("ATENCAO: falha ao escrever em pasta publica.")
            if run_context.get("public_write_error"):
                linhas.append(f"Erro: {run_context['public_write_error']}")
            linhas.append("")

        if run_context["run_status"] == "INCONSISTENT":
            linhas.append("ATENCAO: houve inconsistencias entre a acao e a verificacao.")
            linhas.append("")

        def filtrar_outcome(outcome):
            return [p for p in pessoas if p.get("outcome") == outcome]

        verificados = filtrar_outcome(OUTCOME_VERIFIED_SUCCESS)
        salvos_nao_verificados = filtrar_outcome(OUTCOME_SAVED_NOT_VERIFIED)
        falhas_acao = filtrar_outcome(OUTCOME_FAILED_ACTION)
        falhas_verificacao = filtrar_outcome(OUTCOME_FAILED_VERIFICATION)
        sem_foto = [p for p in pessoas if p.get("no_photo")]
        ignorados = filtrar_outcome(OUTCOME_SKIPPED_ALREADY_EXISTS)
        skipped_dry_run = filtrar_outcome(OUTCOME_SKIPPED_DRY_RUN)
        skipped_no_recipient = filtrar_outcome(OUTCOME_SKIPPED_NO_RECIPIENT)

        if falhas_acao:
            linhas.append("LISTA DE FALHAS NA ACAO:")
            for p in falhas_acao:
                motivo = p.get("errors", {}).get("action_error") or "Falha na acao."
                linhas.append(f" - {p['nome']} ({p['cpf']}): {motivo}")
            linhas.append("")

        if falhas_verificacao:
            linhas.append("LISTA DE FALHAS NA VERIFICACAO:")
            for p in falhas_verificacao:
                motivo = p.get("errors", {}).get("verification_error") or "Falha na verificacao."
                linhas.append(f" - {p['nome']} ({p['cpf']}): {motivo}")
            linhas.append("")

        if salvos_nao_verificados:
            linhas.append("LISTA DE SALVOS (NAO VERIFICADOS):")
            for p in salvos_nao_verificados:
                detalhe = p.get("errors", {}).get("verification_error") or "CPF nao encontrado na lista de rascunhos."
                linhas.append(f" - {p['nome']} ({p['cpf']}): {detalhe}")
            linhas.append("")

        if verificados:
            linhas.append("LISTA DE SUCESSO (VERIFICADO):")
            for p in verificados:
                linhas.append(f" - {p['nome']} ({p['cpf']})")
            linhas.append("")

        if sem_foto:
            linhas.append("LISTA SEM FOTO:")
            for p in sem_foto:
                linhas.append(f" - {p['nome']} ({p['cpf']})")
            linhas.append("")

        if ignorados or skipped_dry_run or skipped_no_recipient:
            linhas.append("LISTA SKIPPED/WARNINGS:")
            for p in ignorados:
                linhas.append(f" - {p['nome']} ({p['cpf']}): {OUTCOME_SKIPPED_ALREADY_EXISTS}")
            for p in skipped_dry_run:
                linhas.append(f" - {p['nome']} ({p['cpf']}): {OUTCOME_SKIPPED_DRY_RUN}")
            for p in skipped_no_recipient:
                linhas.append(f" - {p['nome']} ({p['cpf']}): {OUTCOME_SKIPPED_NO_RECIPIENT}")
            linhas.append("")

        linhas.append(f"Report Path: {report_path}")
        if run_context.get("manifest_path"):
            linhas.append(f"Manifest Path: {run_context['manifest_path']}")

        conteudo = "\n".join(linhas)
        output_manager.write_text(KIND_RELATORIOS, nome_arquivo, conteudo)
        caminho_completo = output_manager.get_preferred_path(KIND_RELATORIOS, nome_arquivo)
        logger.info(f"Relatorio TXT gerado em: {caminho_completo}")
        return caminho_completo

    except Exception as e:
        logger.error(f"Erro ao gerar relatorio TXT: {e}")
        return None


def gerar_relatorio_json(manifest: dict, output_manager: OutputManager):
    """
    Gera um relatorio JSON resumido para padrao operacional.
    """
    try:
        run_context = manifest["run_context"]
        nome_arquivo = f"{_nome_base_execucao(run_context)}__relatorio_resumo.json"
        relatorio = {
            "execution_id": run_context["execution_id"],
            "run_status": run_context.get("run_status"),
            "started_at": run_context.get("started_at"),
            "finished_at": run_context.get("finished_at"),
            "duration_sec": run_context.get("duration_sec"),
            "totals": manifest.get("totals", {}),
            "report_path": run_context.get("report_path"),
            "manifest_path": run_context.get("manifest_path"),
            "public_write_ok": run_context.get("public_write_ok"),
            "public_write_error": run_context.get("public_write_error"),
        }
        output_manager.write_json(KIND_LOGS, nome_arquivo, relatorio)
        caminho = output_manager.get_preferred_path(KIND_LOGS, nome_arquivo)
        logger.info(f"Relatorio JSON gerado em: {caminho}")
        return caminho
    except Exception as e:
        logger.error(f"Erro ao gerar relatorio JSON: {e}")
        return None


def gerar_resumo_execucao_md(manifest: dict, output_manager: OutputManager):
    """
    Gera um resumo executivo em Markdown para padrao operacional.
    """
    try:
        run_context = manifest["run_context"]
        nome_arquivo = f"{_nome_base_execucao(run_context)}__resumo_execucao.md"
        totals = manifest.get("totals", {})

        linhas = []
        linhas.append("# Resumo de Execucao - MetaXg")
        linhas.append("")
        linhas.append(f"- Execution ID: {run_context['execution_id']}")
        linhas.append(f"- Run Status: {run_context.get('run_status')}")
        linhas.append(f"- Started At: {run_context.get('started_at')}")
        if run_context.get("finished_at"):
            linhas.append(f"- Finished At: {run_context.get('finished_at')}")
        if run_context.get("duration_sec") is not None:
            linhas.append(f"- Duration (sec): {run_context.get('duration_sec')}")
        linhas.append("")
        linhas.append("## Totais")
        linhas.append(f"- Detected: {totals.get('detected')}")
        linhas.append(f"- People Total: {totals.get('people_total')}")
        linhas.append(f"- No Photo: {totals.get('no_photo')}")
        linhas.append("")
        linhas.append("## Outcomes")
        for outcome in OUTCOME_ORDER:
            linhas.append(f"- {outcome}: {totals.get('by_outcome', {}).get(outcome, 0)}")
        if totals.get("unknown_outcome"):
            linhas.append(f"- UNKNOWN_OUTCOME: {totals.get('unknown_outcome')}")
        linhas.append("")
        if run_context.get("public_write_ok") is False:
            linhas.append("## Observacoes")
            linhas.append("Falha ao escrever em pasta publica.")
            if run_context.get("public_write_error"):
                linhas.append(f"Erro: {run_context['public_write_error']}")
            linhas.append("")

        conteudo = "\n".join(linhas)
        output_manager.write_text(KIND_LOGS, nome_arquivo, conteudo)
        caminho = output_manager.get_preferred_path(KIND_LOGS, nome_arquivo)
        logger.info(f"Resumo MD gerado em: {caminho}")
        return caminho
    except Exception as e:
        logger.error(f"Erro ao gerar resumo MD: {e}")
        return None


def gerar_diagnostico_ultima_execucao(manifest: dict, output_manager: OutputManager):
    """
    Atualiza diagnostico_ultima_execucao.txt na pasta de logs.
    """
    try:
        run_context = manifest["run_context"]
        totals = manifest.get("totals", {})
        linhas = []
        linhas.append("DIAGNOSTICO ULTIMA EXECUCAO - METAXG")
        linhas.append(f"Execution ID: {run_context.get('execution_id')}")
        linhas.append(f"Run Status: {run_context.get('run_status')}")
        linhas.append(f"Started At: {run_context.get('started_at')}")
        linhas.append(f"Finished At: {run_context.get('finished_at')}")
        if run_context.get("duration_sec") is not None:
            linhas.append(f"Duration (sec): {run_context.get('duration_sec')}")
        linhas.append("")
        linhas.append("Totais:")
        linhas.append(f"- Detected: {totals.get('detected')}")
        linhas.append(f"- People Total: {totals.get('people_total')}")
        linhas.append(f"- No Photo: {totals.get('no_photo')}")
        linhas.append("")
        if run_context.get("manifest_path"):
            linhas.append(f"Manifest Path: {run_context.get('manifest_path')}")
        if run_context.get("report_path"):
            linhas.append(f"Report Path: {run_context.get('report_path')}")
        if run_context.get("public_write_ok") is False:
            linhas.append("")
            linhas.append("ATENCAO: falha ao escrever em pasta publica.")
            if run_context.get("public_write_error"):
                linhas.append(f"Erro: {run_context.get('public_write_error')}")

        conteudo = "\n".join(linhas)
        output_manager.write_text(KIND_LOGS, "diagnostico_ultima_execucao.txt", conteudo)
        caminho = output_manager.get_preferred_path(KIND_LOGS, "diagnostico_ultima_execucao.txt")
        logger.info(f"Diagnostico atualizado em: {caminho}")
        return caminho
    except Exception as e:
        logger.error(f"Erro ao gerar diagnostico: {e}")
        return None


def _sanitizar_nome_arquivo(texto: str) -> str:
    texto = re.sub(r"[^\w\-]+", "_", (texto or "").strip(), flags=re.UNICODE)
    texto = texto.strip("_")
    return texto or "sem_nome"


def _foto_pode_ser_embutida(path: str) -> tuple[bool, str]:
    if not path:
        return False, "caminho_vazio"
    if not os.path.exists(path):
        return False, "arquivo_inexistente"
    if os.path.getsize(path) <= 0:
        return False, "arquivo_vazio"
    try:
        with PILImage.open(path) as img:
            img.verify()
        return True, ""
    except Exception as e:
        return False, str(e)


def gerar_relatorios_erros_pdf(manifest: dict, output_manager: OutputManager) -> list[str]:
    """
    Gera um PDF por pessoa com pendencia de tratativa manual.
    Inclui dados do funcionario, mensagem de erro e foto publica quando disponivel.
    """
    try:
        pessoas = manifest.get("people", [])
        erros = [p for p in pessoas if p.get("outcome") in MANUAL_ERROR_OUTCOMES]
        if not erros:
            return []

        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.platypus import (
            HRFlowable,
            Image,
            Paragraph,
            SimpleDocTemplate,
            Spacer,
            Table,
            TableStyle,
        )

        run_context = manifest["run_context"]
        started_at = datetime.fromisoformat(run_context["started_at"])
        pdf_paths = []

        for index, person in enumerate(erros, start=1):
            dados = person.get("dados_funcionario") or {}
            foto_path = person.get("foto_publica_path") or person.get("foto_path")
            erro_acao = person.get("errors", {}).get("action_error") or ""
            erro_verificacao = person.get("errors", {}).get("verification_error") or ""
            contrato = (person.get("contrato_chave") or "DESCONHECIDO").upper()
            nome_limpo = _sanitizar_nome_arquivo(person.get("nome", "sem_nome"))
            cpf_limpo = _sanitizar_nome_arquivo(person.get("cpf", "sem_cpf"))
            chapa_limpa = _sanitizar_nome_arquivo(dados.get("CHAPA", "sem_chapa"))
            nome_arquivo = (
                f"{started_at.strftime('%Y-%m-%d')}__{started_at.strftime('%Hh%M')}__"
                f"pendencia_{_slug(contrato)}_{index:02d}_{nome_limpo}_{chapa_limpa}.pdf"
            )
            pdf_path = output_manager.get_local_path(KIND_RELATORIOS, nome_arquivo)

            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=A4,
                rightMargin=1.2 * cm,
                leftMargin=1.2 * cm,
                topMargin=1.2 * cm,
                bottomMargin=1.2 * cm,
            )
            styles = getSampleStyleSheet()
            story = []

            title = styles["Title"]
            heading = styles["Heading2"]
            subheading = styles["Heading3"]
            body = styles["BodyText"]
            body.leading = 13

            story.append(Paragraph("Pendencia para Cadastro Manual - MetaXg", title))
            story.append(Spacer(1, 0.25 * cm))
            story.append(Paragraph(f"Funcionario: {person.get('nome', 'Sem nome')}", heading))
            story.append(
                Paragraph(
                    f"Contrato / Frente: <b>{contrato}</b>",
                    body,
                )
            )
            story.append(Paragraph(f"CPF: {person.get('cpf', '')}", body))
            story.append(Paragraph(f"Outcome: {person.get('outcome', '')}", body))
            story.append(Paragraph(f"Execution ID: {run_context['execution_id']}", body))
            story.append(Paragraph(f"Run Status: {run_context.get('run_status')}", body))
            story.append(Paragraph(f"Data/Hora: {started_at.strftime('%d/%m/%Y %H:%M:%S')}", body))
            story.append(Spacer(1, 0.2 * cm))

            if erro_acao:
                story.append(Paragraph(f"Erro de acao: {erro_acao}", body))
            if erro_verificacao:
                story.append(Paragraph(f"Erro de verificacao: {erro_verificacao}", body))
            story.append(Paragraph(f"Sem foto: {'SIM' if person.get('no_photo') else 'NAO'}", body))
            story.append(
                Paragraph(
                    f"Caminho da foto na pasta publica: {foto_path or 'Nao disponivel'}",
                    body,
                )
            )
            if person.get("no_photo"):
                story.append(
                    Paragraph(
                        "Acao necessaria: o robo nao encontrou a foto. A equipe deve acessar a Yube, localizar a foto do colaborador e anexar manualmente no cadastro.",
                        body,
                    )
                )
            story.append(Spacer(1, 0.25 * cm))

            foto_ok, foto_erro = _foto_pode_ser_embutida(foto_path)
            if foto_ok:
                try:
                    story.append(Paragraph("Foto da pasta publica", subheading))
                    story.append(Image(foto_path, width=4.5 * cm, height=6.0 * cm))
                    story.append(Spacer(1, 0.3 * cm))
                except Exception as e:
                    logger.warn("Falha ao embutir foto no PDF de pendencia.", details={"foto": foto_path, "erro": str(e)})
                    story.append(Paragraph(f"Falha ao embutir foto no PDF: {e}", body))
                    story.append(Spacer(1, 0.2 * cm))
            elif foto_path:
                logger.warn("Foto ignorada no PDF de pendencia.", details={"foto": foto_path, "erro": foto_erro})
                story.append(Paragraph(f"Foto nao embutida no PDF: {foto_erro}", body))
                story.append(Spacer(1, 0.2 * cm))

            story.append(Paragraph("Campos do cadastro", subheading))
            rows = [["Campo", "Valor"]]
            for key, value in dados.items():
                texto = "" if value is None else str(value)
                rows.append([str(key), texto if texto else "-"])

            table = Table(rows, colWidths=[5.0 * cm, 11.0 * cm], repeatRows=1)
            table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E6F2")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#8FA1B3")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
                        ("LEADING", (0, 0), (-1, -1), 10),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ]
                )
            )
            story.append(table)
            story.append(Spacer(1, 0.25 * cm))
            story.append(HRFlowable(width="100%", thickness=0.6, color=colors.HexColor("#8FA1B3")))
            story.append(Spacer(1, 0.15 * cm))
            story.append(
                Paragraph(
                    "Este PDF foi gerado para apoiar o cadastro manual quando a automacao nao conclui o processo.",
                    body,
                )
            )

            try:
                doc.build(story)
                with open(pdf_path, "rb") as f:
                    output_manager._write_public_bytes(KIND_RELATORIOS, nome_arquivo, f.read(), started_at)
                pdf_paths.append(output_manager.get_preferred_path(KIND_RELATORIOS, nome_arquivo))
            except Exception as e:
                logger.error(
                    "Falha ao gerar PDF individual de pendencia.",
                    details={"nome": person.get("nome"), "cpf": person.get("cpf"), "erro": str(e)},
                )
                continue

        logger.info("Relatorios PDF individuais de erros gerados", details={"total": len(pdf_paths)})
        return pdf_paths
    except Exception as e:
        logger.error(f"Erro ao gerar relatorios PDF de erros: {e}")
        return []
