#!/usr/bin/env python3
"""
send_email.py

Verstuurt het resultaat van de VLIF4-pipeline via de Resend API,
met alle Excel-bestanden en validatierapporten als bijlage.

Benodigde omgevingsvariabelen:
- RESEND_API_KEY : de API-key van Resend (begint met 're_...')
- NOTIFY_EMAIL   : het e-mailadres van de ontvanger
                   (bij een gratis Resend-account moet dit het adres
                    zijn waarmee je je bij Resend hebt aangemeld)
- RESEND_FROM    : optioneel. Het afzenderadres.
                   Standaard 'VLIF4 Bot <onboarding@resend.dev>'.

Gebruik:
    python send_email.py --attachments-dir attachments --issues-found 0
"""

import argparse
import base64
import os
import sys
from pathlib import Path

import resend


DEFAULT_FROM = "VLIF4 Bot <onboarding@resend.dev>"
MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024  # veilige grens onder Resend's limiet


def collect_attachments(directory: Path):
    """Leest alle bestanden in de map en levert een lijst Resend-bijlagen."""
    attachments = []
    if not directory.is_dir():
        print(f"Waarschuwing: map '{directory}' bestaat niet — geen bijlagen.")
        return attachments

    for path in sorted(directory.iterdir()):
        if not path.is_file():
            continue
        size = path.stat().st_size
        if size > MAX_ATTACHMENT_BYTES:
            print(f"Overgeslagen (te groot, {size} bytes): {path.name}")
            continue
        content = base64.b64encode(path.read_bytes()).decode("ascii")
        attachments.append({"filename": path.name, "content": content})
        print(f"Bijlage toegevoegd: {path.name} ({size} bytes)")

    return attachments


def build_content(issues_found: str, attachments_count: int):
    """Bouwt onderwerp + HTML-body op basis van de issues_found-vlag."""
    has_issues = str(issues_found).strip() not in ("0", "", "false", "False")

    if has_issues:
        subject = "VLIF4-pipeline: validatie vond aandachtspunten"
        html = (
            f"<h2 style='color:#c0392b;margin-bottom:0.4em'>"
            f"Validatie: issues gevonden</h2>"
            f"<p><b>Let op:</b> de validatie heeft één of meer issues "
            f"gedetecteerd. Bekijk de rapporten in bijlage voor details.</p>"
            f"<p>In bijlage vind je {attachments_count} bestand(en): "
            f"de gegenereerde Excel-export(en) en het validatierapport.</p>"
        )
    else:
        subject = "VLIF4-pipeline: alles in orde"
        html = (
            f"<h2 style='color:#27ae60;margin-bottom:0.4em'>"
            f"Validatie: alles OK</h2>"
            f"<p>De validatie is vlekkeloos doorlopen — geen issues gevonden.</p>"
            f"<p>In bijlage vind je {attachments_count} bestand(en): "
            f"de gegenereerde Excel-export(en) en het validatierapport.</p>"
        )

    html += (
        "<p style='color:#666;font-size:0.9em'>"
        "Automatisch verzonden door de VLIF4 GitHub Actions pipeline."
        "</p>"
    )
    return subject, html


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--attachments-dir", required=True)
    parser.add_argument("--issues-found", default="0")
    args = parser.parse_args()

    api_key = os.environ.get("RESEND_API_KEY")
    recipient = os.environ.get("NOTIFY_EMAIL")
    # Oorspronkelijk: sender = os.environ.get("RESEND_FROM", DEFAULT_FROM)
    # Aangepast: een lege RESEND_FROM (secret bestaat niet) valt nu ook terug op de default.
    sender = os.environ.get("RESEND_FROM") or DEFAULT_FROM

    if not all([api_key, recipient]):
        print(
            "Fout: RESEND_API_KEY en NOTIFY_EMAIL moeten beide ingesteld zijn.",
            file=sys.stderr,
        )
        sys.exit(1)

    resend.api_key = api_key

    attachments = collect_attachments(Path(args.attachments_dir))
    subject, html = build_content(args.issues_found, len(attachments))

    # --- Diagnose: laat zien met welke configuratie we vertrekken ---
    print(
        f"Diagnose: key-lengte={len(api_key)}, "
        f"begint met 're_'={api_key.startswith('re_')}"
    )
    print(
        f"Diagnose: from={sender!r}, to={recipient!r}, "
        f"issues_found={args.issues_found!r}, bijlagen={len(attachments)}"
    )

    params = {
        "from": sender,
        "to": [recipient],
        "subject": subject,
        "html": html,
    }
    if attachments:
        params["attachments"] = attachments

    try:
        email = resend.Emails.send(params)
    except Exception as e:
        status = getattr(e, "code", None) or getattr(e, "status_code", "?")
        body = getattr(e, "message", None) or str(e)
        print(f"Resend-fout: {type(e).__name__}", file=sys.stderr)
        print(f"  status: {status}", file=sys.stderr)
        print(f"  body:   {body}", file=sys.stderr)
        sys.exit(1)

    email_id = email.get("id") if isinstance(email, dict) else email
    print(f"Resultaatmail verzonden naar {recipient} (id: {email_id})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
