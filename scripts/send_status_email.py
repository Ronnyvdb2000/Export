#!/usr/bin/env python3
"""
send_status_email.py

Verstuurt een korte statusmail na elke run van de wekelijkse
VLIF4-check, zowel bij een nieuwe versie als wanneer alles bij het
oude blijft. Gebruikt dezelfde SendGrid-secrets als de hoofdpipeline.
"""

import argparse
import os
import sys

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--new-version", choices=["true", "false"], required=True)
    parser.add_argument("--version", default="")
    parser.add_argument("--previous-version", default="")
    args = parser.parse_args()

    api_key = os.environ.get("SENDGRID_API_KEY")
    sender = os.environ.get("SENDGRID_FROM")
    recipient = os.environ.get("NOTIFY_EMAIL")

    if not all([api_key, sender, recipient]):
        print("Fout: SENDGRID_API_KEY, SENDGRID_FROM en NOTIFY_EMAIL moeten allemaal ingesteld zijn.")
        sys.exit(1)

    if args.new_version == "true":
        subject = f"VLIF4-check: nieuwe versie gevonden (blok {args.version})"
        html = (
            f"<p>De wekelijkse VLIF4-check heeft een <b>nieuwe versie</b> gevonden: "
            f"blok {args.version} (voorheen: {args.previous_version or 'onbekend'}).</p>"
            f"<p>De PDF is gedownload en de volledige pipeline "
            f"(build + validatie + export-mail) wordt automatisch gestart.</p>"
            f"<p>Je ontvangt hierover zo dadelijk een aparte mail met het resultaat.</p>"
        )
    else:
        subject = f"VLIF4-check: geen nieuwe versie (nog steeds blok {args.previous_version})"
        html = (
            f"<p>De wekelijkse VLIF4-check is uitgevoerd. Er staat nog steeds "
            f"blok {args.previous_version} online — geen wijziging, geen verdere actie ondernomen.</p>"
        )

    message = Mail(
        from_email=sender,
        to_emails=recipient,
        subject=subject,
        html_content=html,
    )

    sg = SendGridAPIClient(api_key)
    response = sg.send(message)
    print(f"Statusmail verstuurd, statuscode: {response.status_code}")
    if response.status_code >= 300:
        sys.exit(1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
