from datetime import datetime

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

    subject = (
        f"[MetaXg] {run_context['run_status']} | {date_str} | exec={run_context['execution_id']}"
    )

    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2>Relatorio de Execucao - RPA MetaX</h2>
        <p><b>Execution ID:</b> {run_context['execution_id']}</p>
        <p><b>Run Status:</b> {run_context['run_status']}</p>
        <p><b>Data/Hora:</b> {started_at.strftime('%d/%m/%Y %H:%M')}</p>
        <hr>
        <h3>Resumo Geral</h3>
        <ul>
            <li><b>Total Detectado:</b> {totals['detected']}</li>
            <li><b>Total Processado (people):</b> {totals['people_total']}</li>
            <li><b>Sem Foto:</b> {totals['no_photo']}</li>
        </ul>
        <h3>Totais por Outcome</h3>
        <ul>
    """

    for outcome in OUTCOME_ORDER:
        html_body += f"<li><b>{outcome}:</b> {totals['by_outcome'].get(outcome, 0)}</li>"

    if totals.get("unknown_outcome"):
        html_body += f"<li><b>UNKNOWN_OUTCOME:</b> {totals['unknown_outcome']}</li>"

    html_body += "</ul>"

    if run_context.get("public_write_ok") is False:
        html_body += "<p><b>ATENCAO:</b> Falha ao escrever na pasta publica.</p>"
        if run_context.get("public_write_error"):
            html_body += f"<p><b>Erro:</b> {run_context['public_write_error']}</p>"

    if run_context["run_status"] == "INCONSISTENT":
        html_body += "<p><b>ATENCAO:</b> Houve inconsistencias entre acao e verificacao.</p>"

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

    if salvos_nao_verificados:
        html_body += "<h4>Salvos como rascunho (nao verificados):</h4><ul>"
        for p in salvos_nao_verificados:
            motivo = p.get("errors", {}).get("verification_error") or "CPF nao encontrado na lista de rascunhos."
            html_body += f"<li><b>{p['nome']}:</b> {motivo}</li>"
        html_body += "</ul>"

    if verificados:
        html_body += "<h4>Sucessos verificados:</h4><ul>"
        for p in verificados:
            html_body += f"<li><b>{p['nome']}</b></li>"
        html_body += "</ul>"

    if sem_foto:
        html_body += "<h4>Funcionarios sem Foto:</h4><ul>"
        for p in sem_foto:
            html_body += f"<li>{p['nome']}</li>"
        html_body += "</ul>"

    if ignorados or skipped_dry_run or skipped_no_recipient or skipped_email_disabled:
        html_body += "<h4>Skipped/Warnings:</h4><ul>"
        for p in ignorados:
            html_body += f"<li>{p['nome']} - {OUTCOME_SKIPPED_ALREADY_EXISTS}</li>"
        for p in skipped_dry_run:
            html_body += f"<li>{p['nome']} - {OUTCOME_SKIPPED_DRY_RUN}</li>"
        for p in skipped_no_recipient:
            html_body += f"<li>{p['nome']} - {OUTCOME_SKIPPED_NO_RECIPIENT}</li>"
        for p in skipped_email_disabled:
            html_body += f"<li>{p['nome']} - {OUTCOME_SKIPPED_EMAIL_DISABLED}</li>"
        html_body += "</ul>"

    if report_path:
        html_body += f"<p><b>Relatorio TXT:</b> {report_path}</p>"
    if manifest_path:
        try:
            import os

            if os.path.exists(manifest_path):
                html_body += f"<p><b>Manifest (final):</b> {manifest_path}</p>"
            elif partial_manifest_path:
                html_body += f"<p><b>Manifest (parcial):</b> {partial_manifest_path}</p>"
        except Exception:
            if partial_manifest_path:
                html_body += f"<p><b>Manifest (parcial):</b> {partial_manifest_path}</p>"
    elif partial_manifest_path:
        html_body += f"<p><b>Manifest (parcial):</b> {partial_manifest_path}</p>"

    html_body += """
        <br>
        <p><i>Este e um e-mail automatico gerado pelo Robo RPA MetaX.</i></p>
    </body>
    </html>
    """

    attachments = []
    if attachment_paths:
        attachments.extend([p for p in attachment_paths if p])
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
    Envia um relatorio de execucao por e-mail via Outlook Desktop.
    """
    from config import EMAIL_NOTIFICACAO

    if not EMAIL_NOTIFICACAO:
        logger.warn("E-mail de notificacao nao configurado. Relatorio nao sera enviado.")
        return OUTCOME_SKIPPED_NO_RECIPIENT

    try:
        logger.info("Preparando envio de e-mail.", details={"to": EMAIL_NOTIFICACAO})

        import win32com.client  # lazy import for testability

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem

        subject, html_body, attachments = build_email_payload(
            manifest=manifest,
            report_path=report_path,
            manifest_path=manifest_path,
            partial_manifest_path=partial_manifest_path,
            attachment_paths=attachment_paths,
        )

        mail.To = EMAIL_NOTIFICACAO
        mail.Subject = subject
        mail.HTMLBody = html_body

        for path in attachments:
            try:
                mail.Attachments.Add(path)
            except Exception as e:
                logger.warn("Falha ao anexar arquivo.", details={"path": path, "error": str(e)})

        mail.Send()
        logger.info("E-mail de relatorio enviado com sucesso!")
        return "SENT"

    except Exception as e:
        logger.error("Falha ao enviar e-mail.", details={"error": str(e)})
        raise
