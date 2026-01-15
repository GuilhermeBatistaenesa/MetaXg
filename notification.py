import win32com.client
from datetime import datetime
from custom_logger import logger
from config import EMAIL_NOTIFICACAO

def enviar_relatorio_email(stats: dict):
    """
    Envia um relatório de execução por e-mail via Outlook Desktop.
    
    Args:
        stats (dict): Dicionário contendo estatísticas da execução.
            Ex: {
                "total": 10,
                "sucesso": ["João", "Maria"],
                "falha": [{"nome": "José", "motivo": "Erro X"}],
                "sem_foto": ["Pedro"],
                "duplicados": ["Ana"]
            }
    """
    if not EMAIL_NOTIFICACAO:
        logger.warn("E-mail de notificação não configurado. Relatório não será enviado.")
        return

    try:
        logger.info(f"Preparando envio de e-mail para: {EMAIL_NOTIFICACAO}")
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0) # 0 = olMailItem
        
        mail.To = EMAIL_NOTIFICACAO
        mail.Subject = f"Relatório RPA MetaX - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Construção do Corpo do Email em HTML
        html_body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2>Relatório de Execução - RPA MetaX</h2>
            <p><b>Data/Hora:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
            <hr>
            <h3>Resumo Geral</h3>
            <ul>
                <li><b>Total Processado:</b> {stats['total']}</li>
                <li><b>✅ Sucessos:</b> {len(stats['sucesso'])}</li>
                <li><b>⚠️ Ignorados (Sem Foto):</b> {len(stats['sem_foto'])}</li>
                <li><b>⚠️ Duplicados (Já Cadastrados):</b> {len(stats['duplicados'])}</li>
                <li><b>❌ Falhas:</b> {len(stats['falha'])}</li>
            </ul>
            
            <hr>
            
            <h3>Detalhes:</h3>
        """

        if stats['falha']:
            html_body += "<h4>❌ Falhas:</h4><ul>"
            for item in stats['falha']:
                html_body += f"<li><b>{item['nome']}:</b> {item['motivo']}</li>"
            html_body += "</ul>"
            
        if stats['sem_foto']:
            html_body += "<h4>⚠️ Funcionários sem Foto (Não Cadastrados):</h4><ul>"
            for nome in stats['sem_foto']:
                html_body += f"<li>{nome}</li>"
            html_body += "</ul>"

        if stats['duplicados']:
            html_body += "<h4>⚠️ Funcionários já Cadastrados (Ignorados):</h4><ul>"
            for nome in stats['duplicados']:
                html_body += f"<li>{nome}</li>"
            html_body += "</ul>"

        html_body += """
            <br>
            <p><i>Este é um e-mail automático gerado pelo Robô RPA MetaX.</i></p>
        </body>
        </html>
        """
        
        mail.HTMLBody = html_body
        
        mail.Send()
        logger.info("E-mail de relatório enviado com sucesso!")

    except Exception as e:
        logger.error(f"Falha ao enviar e-mail: {e}", details={"error": str(e)})
