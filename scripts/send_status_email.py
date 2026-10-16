#!/usr/bin/env python3
"""
send_status_email.py

Verstuurt een korte statusmail na elke run van de wekelijkse
VLIF4-check, zowel bij een nieuwe versie als wanneer alles bij het
oude blijft. Gebruikt de Resend API (in plaats van SendGrid).

Benodigde omgevingsvariabelen:
- RESEND_API_KEY : de API-key van Resend (begint met 're_...')
- NOTIFY_EMAIL   : het e-mailadres van de ontvanger
                   (bij een gratis Resend-account moet dit het adres
                    zijn waarmee je je bij Resend hebt aangemeld)
- RESEND_FROM    : optioneel. Het afzenderadres.
                   Standaard 'VLIF4 Bot <onboarding@resend.dev>'.
"""

import argparse
import os
import sys

import resend


DEFAULT_FROM = "VLIF4 Bot <onboarding@resend.dev>"


def build_content(new_version: str, version: str, previous_version: str):
    """Bouwt onderwerp + HTML-body op basis van de argumenten."""
    if new_version == "true":
        subject = f"VLIF4-check: nieuwe versie gevonden (blok {version})"
        html = (
            f"<p>De wekelijkse VLIF4-check heeft een <b>nieuwe versie</b> gevonden: "
            f"blok {version} (voorheen: {previous_version or 'onbekend'}).</p>"
            f"<p>De PDF is gedownload en de volledige pipeline "
            f"(build + validatie + export-mail) wordt automatisch gestart.</p>"
            f"<p>Je ontvangt hierover zo dadelijk een aparte mail met het resultaat.</p>"
        )
    else:
        subject = (
            f"VLIF4-check: geen nieuwe versie "
            f"(nog steeds blok {previous_version or 'onbekend'})"
        )
        html = (
            f"<p>De wekelijkse VLIF4-check is uitgevoerd. Er staat nog steeds "
            f"blok {previous_version or 'onbekend'} online — geen wijziging, "
            f"geen verdere actie ondernomen.</p>"
        )
    return subject, html


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--new-version", choices=["true", "false"], required=True)
    parser.add_argument("--version", default="")
    parser.add_argument("--previous-version", default="")
    args = parser.parse_args()

    api_key = os.environ.get("RESEND_API_KEY")
    recipient = os.environ.get("NOTIFY_EMAIL")
    sender = os.environ.get("RESEND_FROM", DEFAULT_FROM)

    if not all([api_key, recipient]):
        print(
            "Fout: RESEND_API_KEY en NOTIFY_EMAIL moeten beide ingesteld zijn.",
            file=sys.stderr,
        )
        sys.exit(1)

    resend.api_key = api_key

    subject, html = build_content(
        args.new_version, args.version, args.previous_version
    )

    # --- Diagnose: laat zien met welke configuratie we vertrekken ---
    print(
        f"Diagnose: key-lengte={len(api_key)}, "
        f"begint met 're_'={api_key.startswith('re_')}"
    )
    print(f"Diagnose: from={sender!r}, to={recipient!r}")

    params: resend.Emails.SendParams = {
        "from": sender,
        "to": [recipient],
        "subject": subject,
        "html": html,
    }

    try:
        email = resend.Emails.send(params)
    except Exception as e:
        # Probeer zoveel mogelijk detail uit de exception te halen
        status = getattr(e, "code", None) or getattr(e, "status_code", "?")
        body = getattr(e, "message", None) or str(e)
        print(f"Resend-fout: {type(e).__name__}", file=sys.stderr)
        print(f"  status: {status}", file=sys.stderr)
        print(f"  body:   {body}", file=sys.stderr)
        sys.exit(1)

    # De SDK geeft normaal een dict terug met minstens een 'id'
    email_id = email.get("id") if isinstance(email, dict) else email
    print(f"Statusmail verstuurd naar {recipient} (id: {email_id})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
