from datetime import datetime
from email.message import EmailMessage
from email.utils import formatdate
import mimetypes
import os
import smtplib
import zipfile

from custom_logger import logger
from outcomes import (
    OUTCOME_FAILED_ACTION,
    OUTCOME_FAILED_VERIFICATION,
    OUTCOME_SAVED_NOT_VERIFIED,
    OUTCOME_VERIFIED_SUCCESS,
    OUTCOME_SKIPPED_ALREADY_EXISTS,
    OUTCOME_SKIPPED_DRY_RUN,
    OUTCOME_SKIPPED_NO_RECIPIENT,
    OUTCOME_SKIPPED_EMAIL_DISABLED,
    OUTCOME_ORDER,
)


def _parse_recipients(raw_value: str) -> list[str]:
    normalized = (raw_value or "").replace(";", ",")
    recipients = []
    for item in normalized.split(","):
        email = item.strip().strip('"').strip("'")
        if email and email not in recipients:
            recipients.append(email)
    return recipients


def _build_email_message(
    subject: str,
    html_body: str,
    recipients: list[str],
    attachments: list[str],
    sender: str | None = None,
) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["To"] = ", ".join(recipients)
    if sender:
        msg["From"] = sender
    msg["Date"] = formatdate(localtime=True)
    msg.set_content("Seu cliente de e-mail nao suporta HTML.")
    msg.add_alternative(html_body, subtype="html")

    for path in attachments:
        if not path or not os.path.exists(path):
            logger.warn("Anexo nao encontrado para envio.", details={"path": path})
            continue
        mime_type, _ = mimetypes.guess_type(path)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        with open(path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(path),
            )
    return msg


def _attachment_size(path: str) -> int:
    try:
        return os.path.getsize(path)
    except Exception:
        return 0


def _prepare_single_email_attachments(
    manifest: dict,
    attachments: list[str],
    max_attachments_mb: int,
) -> tuple[list[str], str | None]:
    arquivos = []
    for path in attachments or []:
        if path and path not in arquivos:
            arquivos.append(path)

    if not arquivos:
        return [], None

    base_attachments = []
    pdf_attachments = []
    outros = []
    for path in arquivos:
        ext = os.path.splitext(path)[1].lower()
        if ext in {".txt", ".json"}:
            base_attachments.append(path)
        elif ext == ".pdf":
            pdf_attachments.append(path)
        else:
            outros.append(path)

    final_attachments = list(base_attachments) + list(outros)
    notice = None

    if pdf_attachments:
        run_context = manifest.get("run_context", {})
        execution_id = run_context.get("execution_id", "sem_execucao")
        output_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(output_dir, exist_ok=True)
        zip_path = os.path.join(output_dir, f"pendencias_email_{execution_id}.zip")

        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
            for path in pdf_attachments:
                if os.path.exists(path):
                    zipf.write(path, arcname=os.path.basename(path))

        limite_bytes = max(1, max_attachments_mb) * 1024 * 1024
        total_com_zip = sum(_attachment_size(path) for path in final_attachments) + _attachment_size(zip_path)
        if total_com_zip <= limite_bytes:
            final_attachments.append(zip_path)
            notice = f"PDFs de pendencia enviados em um unico ZIP: {zip_path}"
        else:
            notice = (
                "PDFs de pendencia consolidados em ZIP, mas nao anexados para evitar rejeicao por tamanho. "
                f"Arquivo disponivel em: {zip_path}"
            )

    return final_attachments, notice


def _enviar_via_outlook(
    recipients_raw: str,
    subject: str,
    html_body: str,
    attachments: list[str],
) -> str:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        # Tenta usar o Outlook já aberto antes de criar nova instância
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception:
                outlook = win32com.client.DispatchEx("Outlook.Application")

        try:
            outlook.GetNamespace("MAPI").Logon("", "", False, False)
        except Exception:
            pass

        mail = outlook.CreateItem(0)
        mail.To = recipients_raw
        mail.Subject = subject
        mail.HTMLBody = html_body

        for path in attachments:
            if not path or not os.path.exists(path):
                logger.warn("Anexo nao encontrado para Outlook.", details={"path": path})
                continue
            try:
                mail.Attachments.Add(path)
            except Exception as e:
                logger.warn("Falha ao anexar arquivo no Outlook.", details={"path": path, "error": str(e)})

        mail.Send()
        logger.info("E-mail de relatorio enviado com sucesso via Outlook.")
        return "SENT_OUTLOOK"
    finally:
        pythoncom.CoUninitialize()


def _enviar_via_smtp(
    recipients: list[str],
    subject: str,
    html_body: str,
    attachments: list[str],
) -> str:
    from config import (
        EMAIL_REMETENTE,
        EMAIL_SMTP_HOST,
        EMAIL_SMTP_PASSWORD,
        EMAIL_SMTP_PORT,
        EMAIL_SMTP_USER,
        EMAIL_SMTP_USE_SSL,
        EMAIL_SMTP_USE_TLS,
    )

    if not EMAIL_SMTP_HOST:
        raise RuntimeError("SMTP nao configurado: EMAIL_SMTP_HOST vazio.")

    sender = EMAIL_REMETENTE or EMAIL_SMTP_USER
    if not sender:
        raise RuntimeError("SMTP nao configurado: defina EMAIL_REMETENTE ou EMAIL_SMTP_USER.")

    msg = _build_email_message(subject, html_body, recipients, attachments, sender=sender)

    if EMAIL_SMTP_USE_SSL:
        server = smtplib.SMTP_SSL(EMAIL_SMTP_HOST, EMAIL_SMTP_PORT, timeout=30)
    else:
        server = smtplib.SMTP(EMAIL_SMTP_HOST, EMAIL_SMTP_PORT, timeout=30)

    try:
        server.ehlo()
        if EMAIL_SMTP_USE_TLS and not EMAIL_SMTP_USE_SSL:
            server.starttls()
            server.ehlo()
        if EMAIL_SMTP_USER:
            server.login(EMAIL_SMTP_USER, EMAIL_SMTP_PASSWORD)
        server.send_message(msg)
    finally:
        try:
            server.quit()
        except Exception:
            pass

    logger.info("E-mail de relatorio enviado com sucesso via SMTP.", details={"host": EMAIL_SMTP_HOST, "port": EMAIL_SMTP_PORT})
    return "SENT_SMTP"


def _salvar_email_eml(
    manifest: dict,
    recipients: list[str],
    subject: str,
    html_body: str,
    attachments: list[str],
) -> str:
    run_context = manifest.get("run_context", {})
    execution_id = run_context.get("execution_id", "sem_execucao")
    started_at = run_context.get("started_at")
    try:
        dt = datetime.fromisoformat(started_at) if started_at else datetime.now()
    except Exception:
        dt = datetime.now()

    filename = f"email_fallback_{dt.strftime('%Y-%m-%d_%H-%M-%S')}__{execution_id}.eml"
    output_dir = os.path.join(os.getcwd(), "logs")
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, filename)

    msg = _build_email_message(subject, html_body, recipients, attachments)
    with open(path, "wb") as f:
        f.write(bytes(msg))

    logger.warn("E-mail salvo em .eml para envio manual.", details={"path": path})
    return path


def build_email_payload(
    manifest: dict,
    report_path: str | None = None,
    manifest_path: str | None = None,
    partial_manifest_path: str | None = None,
    attachment_paths: list[str] | None = None,
) -> tuple[str, str, list[str]]:
    run_context = manifest["run_context"]
    totals = manifest["totals"]
    pessoas = manifest["people"]
    started_at = datetime.fromisoformat(run_context["started_at"])
    date_str = started_at.strftime("%Y-%m-%d")

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
    skipped_email_disabled = filtrar_outcome(OUTCOME_SKIPPED_EMAIL_DISABLED)
    pendencias_manuais = falhas_acao + falhas_verificacao + salvos_nao_verificados

    n_pendencias = len(pendencias_manuais)
    n_ok = len(verificados)
    n_sem_foto = totals.get("no_photo", 0)

    if n_pendencias > 0:
        status_subject = f"{n_pendencias} pendência(s) - verificar manualmente"
    elif n_ok > 0:
        status_subject = f"{n_ok} cadastro(s) confirmado(s)"
    else:
        status_subject = "sem novos cadastros"

    subject = f"[RPA MetaX] {started_at.strftime('%d/%m/%Y')} - {status_subject}"

    CORES = {
        "vermelho": "#c0392b",
        "amarelo": "#e67e22",
        "verde": "#27ae60",
        "cinza": "#7f8c8d",
        "azul": "#2980b9",
    }

    def secao(titulo, cor, itens_html):
        return f"""
        <div style="margin:16px 0;padding:12px 16px;border-left:4px solid {cor};background:#fafafa;">
            <b style="color:{cor}">{titulo}</b>
            {itens_html}
        </div>"""

    html_body = f"""
    <html>
    <body style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:700px;">
        <h2 style="margin-bottom:4px;">Relatório RPA MetaX — {started_at.strftime('%d/%m/%Y %H:%M')}</h2>
        <p style="color:#555;margin-top:0;">
            Total detectado no RM: <b>{totals['detected']}</b> &nbsp;|&nbsp;
            Cadastrados com sucesso: <b style="color:{CORES['verde']}">{n_ok}</b> &nbsp;|&nbsp;
            Pendências: <b style="color:{CORES['vermelho']}">{n_pendencias}</b> &nbsp;|&nbsp;
            Sem foto: <b style="color:{CORES['amarelo']}">{n_sem_foto}</b>
        </p>
        <hr style="border:none;border-top:1px solid #ddd;">
    """

    if run_context.get("public_write_ok") is False:
        html_body += secao(
            "⚠ Falha ao gravar na pasta compartilhada",
            CORES["vermelho"],
            f"<p style='margin:4px 0'>{run_context.get('public_write_error','Erro desconhecido')}</p>",
        )

    if falhas_acao:
        itens = "<ul style='margin:6px 0'>"
        for p in falhas_acao:
            motivo = p.get("errors", {}).get("action_error") or "Falha no cadastro."
            itens += f"<li><b>{p['nome']}</b> — {motivo}</li>"
        itens += "</ul>"
        itens += "<p style='margin:4px 0'>👉 Estes colaboradores <b>não foram cadastrados</b>. Faça o cadastro manual no MetaX.</p>"
        html_body += secao(f"❌ Falha no cadastro ({len(falhas_acao)})", CORES["vermelho"], itens)

    if falhas_verificacao:
        itens = "<ul style='margin:6px 0'>"
        for p in falhas_verificacao:
            motivo = p.get("errors", {}).get("verification_error") or "Falha na verificacao."
            itens += f"<li><b>{p['nome']}</b> — {motivo}</li>"
        itens += "</ul>"
        itens += "<p style='margin:4px 0'>👉 O cadastro foi feito mas o robô não conseguiu confirmar. Verifique no MetaX se o rascunho existe.</p>"
        html_body += secao(f"⚠ Verificação com erro ({len(falhas_verificacao)})", CORES["amarelo"], itens)

    if salvos_nao_verificados:
        itens = "<ul style='margin:6px 0'>"
        for p in salvos_nao_verificados:
            itens += f"<li><b>{p['nome']}</b></li>"
        itens += "</ul>"
        itens += "<p style='margin:4px 0'>👉 Foram salvos como rascunho mas o robô não confirmou na lista. Confira no MetaX se os rascunhos aparecem corretamente.</p>"
        html_body += secao(f"🔄 Rascunho salvo — aguardando confirmação ({len(salvos_nao_verificados)})", CORES["azul"], itens)

    if sem_foto:
        itens = "<ul style='margin:6px 0'>"
        for p in sem_foto:
            itens += f"<li><b>{p['nome']}</b></li>"
        itens += "</ul>"
        itens += "<p style='margin:4px 0'>👉 Foto não encontrada automaticamente. Acesse a Yube, baixe a foto e inclua manualmente no cadastro do MetaX.</p>"
        html_body += secao(f"📷 Sem foto ({len(sem_foto)})", CORES["amarelo"], itens)

    if verificados:
        itens = "<ul style='margin:6px 0'>"
        for p in verificados:
            itens += f"<li>{p['nome']}</li>"
        itens += "</ul>"
        html_body += secao(f"✅ Cadastrados com sucesso ({len(verificados)})", CORES["verde"], itens)

    if ignorados:
        itens = "<ul style='margin:6px 0'>"
        for p in ignorados:
            itens += f"<li>{p['nome']}</li>"
        itens += "</ul>"
        html_body += secao(f"⏭ Já existiam no sistema — ignorados ({len(ignorados)})", CORES["cinza"], itens)

    skipped_outros = skipped_dry_run + skipped_no_recipient + skipped_email_disabled
    if skipped_outros:
        itens = "<ul style='margin:6px 0'>"
        for p in skipped_outros:
            itens += f"<li>{p['nome']}</li>"
        itens += "</ul>"
        html_body += secao(f"⏭ Outros ignorados ({len(skipped_outros)})", CORES["cinza"], itens)

    if pendencias_manuais:
        html_body += "<p style='color:#555;font-size:13px;margin-top:16px'>📎 Seguem em anexo os PDFs individuais das pendências com os dados do cadastro.</p>"

    if report_path:
        html_body += f"<p style='color:#555;font-size:12px'>Relatório completo: {report_path}</p>"

    html_body += """
        <hr style="border:none;border-top:1px solid #ddd;margin-top:20px">
        <p style="color:#aaa;font-size:12px;margin-top:8px"><i>E-mail automático gerado pelo Robô RPA MetaX.</i></p>
    </body>
    </html>
    """

    attachments = []
    if attachment_paths:
        for path in attachment_paths:
            if path and path not in attachments:
                attachments.append(path)
    else:
        if report_path:
            attachments.append(report_path)
        if manifest_path:
            attachments.append(manifest_path)
        elif partial_manifest_path:
            attachments.append(partial_manifest_path)

    return subject, html_body, attachments


def enviar_relatorio_email(
    manifest: dict,
    report_path: str | None = None,
    manifest_path: str | None = None,
    partial_manifest_path: str | None = None,
    attachment_paths: list[str] | None = None,
) -> str:
    """
    Envia um relatorio de execucao por e-mail via Outlook Desktop com fallback para SMTP.
    """
    from config import EMAIL_MAX_ATTACHMENTS_MB, EMAIL_NOTIFICACAO

    if not EMAIL_NOTIFICACAO:
        logger.warn("E-mail de notificacao nao configurado. Relatorio nao sera enviado.")
        return OUTCOME_SKIPPED_NO_RECIPIENT

    recipients = _parse_recipients(EMAIL_NOTIFICACAO)
    if not recipients:
        logger.warn("Nenhum destinatario valido apos normalizacao do campo EMAIL_NOTIFICACAO.")
        return OUTCOME_SKIPPED_NO_RECIPIENT

    try:
        logger.info("Preparando envio de e-mail.", details={"to": EMAIL_NOTIFICACAO})

        subject, html_body, attachments = build_email_payload(
            manifest=manifest,
            report_path=report_path,
            manifest_path=manifest_path,
            partial_manifest_path=partial_manifest_path,
            attachment_paths=attachment_paths,
        )
        attachments, attachment_notice = _prepare_single_email_attachments(
            manifest=manifest,
            attachments=attachments,
            max_attachments_mb=EMAIL_MAX_ATTACHMENTS_MB,
        )
        if attachment_notice:
            html_body = html_body.replace(
                "</body>",
                f"<p><b>Anexos consolidados:</b> {attachment_notice}</p></body>",
            )

        erros = []

        try:
            _enviar_via_outlook(EMAIL_NOTIFICACAO, subject, html_body, attachments)
            return "SENT_OUTLOOK"
        except Exception as e:
            erros.append(f"Outlook: {e}")
            logger.warn("Envio via Outlook falhou. Tentando SMTP...", details={"error": str(e)})

        try:
            _enviar_via_smtp(recipients, subject, html_body, attachments)
            return "SENT_SMTP"
        except Exception as e:
            erros.append(f"SMTP: {e}")
            logger.warn("Envio via SMTP falhou.", details={"error": str(e)})

        eml_path = _salvar_email_eml(manifest, recipients, subject, html_body, attachments)
        raise RuntimeError(" | ".join(erros + [f"Fallback .eml salvo em {eml_path}"]))

    except Exception as e:
        logger.error("Falha ao enviar e-mail.", details={"error": str(e)})
        raise
