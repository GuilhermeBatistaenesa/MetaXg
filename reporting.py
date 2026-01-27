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
        linhas.append(f"- Ignorados: {totals['ignored']}")
        linhas.append(f"- Processados com Sucesso (verificado): {totals['processed_success']}")
        linhas.append(f"- Falhas: {totals['failed']}")
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

        def filtrar(status):
            return [p for p in pessoas if p["status"] == status]

        ignorados = filtrar("IGNORED")
        sucessos = filtrar("SUCCESS")
        falhas = filtrar("FAILED")
        sem_foto = [p for p in pessoas if p.get("no_photo")]

        if ignorados:
            linhas.append("LISTA DE IGNORADOS:")
            for p in ignorados:
                detalhe = p.get("verification_detail", "").strip()
                linha = f" - {p['nome']} ({p['cpf']})"
                if detalhe:
                    linha += f": {detalhe}"
                linhas.append(linha)
            linhas.append("")

        if sucessos:
            linhas.append("LISTA DE SUCESSO (VERIFICADO):")
            for p in sucessos:
                linhas.append(f" - {p['nome']} ({p['cpf']})")
            linhas.append("")

        if sem_foto:
            linhas.append("LISTA SEM FOTO:")
            for p in sem_foto:
                linhas.append(f" - {p['nome']} ({p['cpf']})")
            linhas.append("")

        if falhas:
            linhas.append("LISTA DE FALHAS:")
            for p in falhas:
                motivo = p.get("error") or p.get("verification_detail") or "Motivo nao especificado"
                linhas.append(f" - {p['nome']} ({p['cpf']}): {motivo}")
            linhas.append("")

        conteudo = "\n".join(linhas)
        caminho_completo = output_manager.write_text(KIND_RELATORIOS, nome_arquivo, conteudo)
        logger.info(f"Relatorio TXT gerado em: {caminho_completo}")
        return caminho_completo

    except Exception as e:
        logger.error(f"Erro ao gerar relatorio TXT: {e}")
        return None
