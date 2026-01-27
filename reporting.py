from datetime import datetime

from custom_logger import logger
from output_manager import KIND_RELATORIOS, OutputManager


def gerar_relatorio_txt(manifest: dict, output_manager: OutputManager):
    """
    Gera um relatorio TXT detalhado da execucao a partir do manifest.
    """
    try:
        started_at = datetime.fromisoformat(manifest["started_at"])
        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        nome_arquivo = f"relatorio_execucao_{data_str}__{manifest['execution_id']}.txt"

        totals = manifest["totals"]
        pessoas = manifest["people"]

        linhas = []
        linhas.append("RELATORIO DE EXECUCAO RPA METAX")
        linhas.append(f"Data: {started_at.strftime('%d/%m/%Y %H:%M:%S')}")
        linhas.append(f"Run Status: {manifest['run_status']}")
        linhas.append("=" * 50)
        linhas.append("")
        linhas.append("RESUMO GERAL:")
        linhas.append(f"- Total Detectado no RM: {totals['detected']}")
        linhas.append(f"- Salvos como rascunho (total): {totals['action_saved']}")
        linhas.append(f"- Sucessos verificados: {totals['verified_success']}")
        linhas.append(f"- Salvos como rascunho (nao verificados): {totals['saved_not_verified']}")
        linhas.append(f"- Falhas na acao: {totals['failed_action']}")
        linhas.append(f"- Falhas na verificacao: {totals['failed_verification']}")
        linhas.append(f"- Sem Foto: {totals['no_photo']}")
        linhas.append("")

        if manifest.get("public_write_ok") is False:
            linhas.append("ATENCAO: falha ao escrever em pasta publica.")
            if manifest.get("public_write_error"):
                linhas.append(f"Erro: {manifest['public_write_error']}")
            linhas.append("")

        if manifest["run_status"] == "INCONSISTENT":
            linhas.append("ATENCAO: houve inconsistencias entre a acao e a verificacao.")
            linhas.append("")

        def filtrar_outcome(outcome):
            return [p for p in pessoas if p.get("outcome") == outcome]

        verificados = filtrar_outcome("VERIFIED_SUCCESS")
        salvos_nao_verificados = filtrar_outcome("SAVED_NOT_VERIFIED")
        falhas_acao = filtrar_outcome("FAILED_ACTION")
        falhas_verificacao = filtrar_outcome("FAILED_VERIFICATION")
        sem_foto = [p for p in pessoas if p.get("no_photo")]

        if verificados:
            linhas.append("LISTA DE SUCESSO (VERIFICADO):")
            for p in verificados:
                linhas.append(f" - {p['nome']} ({p['cpf']})")
            linhas.append("")

        if salvos_nao_verificados:
            linhas.append("LISTA DE SALVOS (NAO VERIFICADOS):")
            for p in salvos_nao_verificados:
                detalhe = p.get("errors", {}).get("verification_error") or "CPF nao encontrado na lista de rascunhos."
                linhas.append(f" - {p['nome']} ({p['cpf']}): {detalhe}")
            linhas.append("")

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

        if sem_foto:
            linhas.append("LISTA SEM FOTO:")
            for p in sem_foto:
                linhas.append(f" - {p['nome']} ({p['cpf']})")
            linhas.append("")

        conteudo = "\n".join(linhas)
        caminho_completo = output_manager.write_text(KIND_RELATORIOS, nome_arquivo, conteudo)
        logger.info(f"Relatorio TXT gerado em: {caminho_completo}")
        return caminho_completo

    except Exception as e:
        logger.error(f"Erro ao gerar relatorio TXT: {e}")
        return None
