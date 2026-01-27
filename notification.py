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
                <li><b>Sucessos verificados:</b> {totals['processed_success']}</li>
                <li><b>Ignorados:</b> {totals['ignored']}</li>
                <li><b>Falhas:</b> {totals['failed']}</li>
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
        falhas = [p for p in pessoas if p["status"] == "FAILED"]
        ignorados = [p for p in pessoas if p["status"] == "IGNORED"]
        sem_foto = [p for p in pessoas if p.get("no_photo")]

        if falhas:
            html_body += "<h4>Falhas:</h4><ul>"
            for p in falhas:
                motivo = p.get("error") or p.get("verification_detail") or "Motivo nao especificado"
                html_body += f"<li><b>{p['nome']}:</b> {motivo}</li>"
            html_body += "</ul>"

        if sem_foto:
            html_body += "<h4>Funcionarios sem Foto:</h4><ul>"
            for p in sem_foto:
                html_body += f"<li>{p['nome']}</li>"
            html_body += "</ul>"

        if ignorados:
            html_body += "<h4>Ignorados (ja cadastrados):</h4><ul>"
            for p in ignorados:
                html_body += f"<li>{p['nome']}</li>"
            html_body += "</ul>"

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
