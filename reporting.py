from datetime import datetime

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
        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        nome_arquivo = f"relatorio_execucao_{data_str}__{run_context['execution_id']}.txt"

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
        caminho_completo = output_manager.write_text(KIND_RELATORIOS, nome_arquivo, conteudo)
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
        started_at = datetime.fromisoformat(run_context["started_at"])
        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        nome_arquivo = f"relatorio_{data_str}__{run_context['execution_id']}.json"
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
        caminho = output_manager.write_json(KIND_LOGS, nome_arquivo, relatorio)
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
        started_at = datetime.fromisoformat(run_context["started_at"])
        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        nome_arquivo = f"resumo_execucao_{data_str}__{run_context['execution_id']}.md"
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
        caminho = output_manager.write_text(KIND_LOGS, nome_arquivo, conteudo)
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
        caminho = output_manager.write_text(KIND_LOGS, "diagnostico_ultima_execucao.txt", conteudo)
        logger.info(f"Diagnostico atualizado em: {caminho}")
        return caminho
    except Exception as e:
        logger.error(f"Erro ao gerar diagnostico: {e}")
        return None
