import win32com.client
from datetime import datetime
from custom_logger import logger
from config import EMAIL_NOTIFICACAO


def enviar_relatorio_email(manifest: dict):
    """
    Envia um relatorio de execucao por e-mail via Outlook Desktop.
    """
    if not EMAIL_NOTIFICACAO:
        logger.warn("E-mail de notificacao nao configurado. Relatorio nao sera enviado.")
        return

    try:
        logger.info(f"Preparando envio de e-mail para: {EMAIL_NOTIFICACAO}")

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem

        started_at = datetime.fromisoformat(manifest["started_at"])
        totals = manifest["totals"]

        mail.To = EMAIL_NOTIFICACAO
        mail.Subject = f"Relatorio RPA MetaX - {started_at.strftime('%d/%m/%Y %H:%M')}"

        html_body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2>Relatorio de Execucao - RPA MetaX</h2>
            <p><b>Data/Hora:</b> {started_at.strftime('%d/%m/%Y %H:%M')}</p>
            <p><b>Run Status:</b> {manifest['run_status']}</p>
            <hr>
            <h3>Resumo Geral</h3>
            <ul>
                <li><b>Total Detectado:</b> {totals['detected']}</li>
                <li><b>Salvos como rascunho (total):</b> {totals['action_saved']}</li>
                <li><b>Sucessos verificados:</b> {totals['verified_success']}</li>
                <li><b>Salvos como rascunho (nao verificados):</b> {totals['saved_not_verified']}</li>
                <li><b>Falhas na acao:</b> {totals['failed_action']}</li>
                <li><b>Falhas na verificacao:</b> {totals['failed_verification']}</li>
                <li><b>Sem Foto:</b> {totals['no_photo']}</li>
            </ul>
        """

        if manifest.get("public_write_ok") is False:
            html_body += "<p><b>ATENCAO:</b> Falha ao escrever na pasta publica.</p>"
            if manifest.get("public_write_error"):
                html_body += f"<p><b>Erro:</b> {manifest['public_write_error']}</p>"

        if manifest["run_status"] == "INCONSISTENT":
            html_body += "<p><b>ATENCAO:</b> Houve inconsistencias entre acao e verificacao.</p>"

        pessoas = manifest["people"]
        verificados = [p for p in pessoas if p.get("outcome") == "VERIFIED_SUCCESS"]
        salvos_nao_verificados = [p for p in pessoas if p.get("outcome") == "SAVED_NOT_VERIFIED"]
        falhas_acao = [p for p in pessoas if p.get("outcome") == "FAILED_ACTION"]
        falhas_verificacao = [p for p in pessoas if p.get("outcome") == "FAILED_VERIFICATION"]
        sem_foto = [p for p in pessoas if p.get("no_photo")]

        if verificados:
            html_body += "<h4>Sucessos verificados:</h4><ul>"
            for p in verificados:
                html_body += f"<li><b>{p['nome']}</b></li>"
            html_body += "</ul>"

        if salvos_nao_verificados:
            html_body += "<h4>Salvos como rascunho (nao verificados):</h4><ul>"
            for p in salvos_nao_verificados:
                motivo = p.get("errors", {}).get("verification_error") or "CPF nao encontrado na lista de rascunhos."
                html_body += f"<li><b>{p['nome']}:</b> {motivo}</li>"
            html_body += "</ul>"

        if falhas_acao:
            html_body += "<h4>Falhas na acao:</h4><ul>"
            for p in falhas_acao:
                motivo = p.get("errors", {}).get("action_error") or "Falha na acao."
                html_body += f"<li><b>{p['nome']}:</b> {motivo}</li>"
            html_body += "</ul>"

        if falhas_verificacao:
            html_body += "<h4>Falhas na verificacao:</h4><ul>"
            for p in falhas_verificacao:
                motivo = p.get("errors", {}).get("verification_error") or "Falha na verificacao."
                html_body += f"<li><b>{p['nome']}:</b> {motivo}</li>"
            html_body += "</ul>"

        if sem_foto:
            html_body += "<h4>Funcionarios sem Foto:</h4><ul>"
            for p in sem_foto:
                html_body += f"<li>{p['nome']}</li>"
            html_body += "</ul>"

        # Ignorados nao entram mais como sucesso nem falha (mantido no manifest com outcome FAILED_ACTION)

        html_body += """
            <br>
            <p><i>Este e um e-mail automatico gerado pelo Robo RPA MetaX.</i></p>
        </body>
        </html>
        """

        mail.HTMLBody = html_body
        mail.Send()
        logger.info("E-mail de relatorio enviado com sucesso!")

    except Exception as e:
        logger.error(f"Falha ao enviar e-mail: {e}", details={"error": str(e)})
