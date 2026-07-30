"""
send_email.py

Vierde stap in de VLIF4-pipeline: stuurt het Excel-bestand en het
validatierapport als bijlage via SendGrid.

Benodigde environment variables (als GitHub secrets ingesteld):
  SENDGRID_API_KEY   - API-key van je SendGrid-account
  SENDGRID_FROM      - geverifieerd afzenderadres (single sender verification)
  NOTIFY_EMAIL       - ontvanger; komt normaal uit de workflow_dispatch input,
                       met een secret als fallback als er geen input is opgegeven
"""

import argparse
import base64
import os
import sys

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition


def build_attachment(path: str, mime_type: str) -> Attachment:
    with open(path, "rb") as f:
        data = f.read()
    encoded = base64.b64encode(data).decode()
    return Attachment(
        FileContent(encoded),
        FileName(os.path.basename(path)),
        FileType(mime_type),
        Disposition("attachment"),
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", required=True, help="pad naar het finale Excel-bestand")
    parser.add_argument("--report", required=True, help="pad naar validation_report.xlsx")
    parser.add_argument("--issues-found", type=int, default=0,
                         help="1 als validate_output.py issues vond, anders 0")
    args = parser.parse_args()

    api_key = os.environ.get("SENDGRID_API_KEY")
    sender = os.environ.get("SENDGRID_FROM")
    recipient = os.environ.get("NOTIFY_EMAIL")

    if not all([api_key, sender, recipient]):
        print("Fout: SENDGRID_API_KEY, SENDGRID_FROM en NOTIFY_EMAIL moeten allemaal ingesteld zijn.")
        sys.exit(1)

    status = "⚠️ Validatieproblemen gevonden — zie rapport" if args.issues_found else "✅ Validatie geslaagd, geen issues"

    message = Mail(
        from_email=sender,
        to_emails=recipient,
        subject=f"VLIF4-export klaar — {status}",
        html_content=(
            f"<p>De VLIF4-pipeline is uitgevoerd.</p>"
            f"<p><b>Status:</b> {status}</p>"
            f"<p>Het Excel-bestand en het validatierapport zitten als bijlage.</p>"
        ),
    )

    message.attachment = [
        build_attachment(args.excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        build_attachment(args.report, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ]

    sg = SendGridAPIClient(api_key)
    response = sg.send(message)
    print(f"Mail verstuurd, statuscode: {response.status_code}")

    if response.status_code >= 300:
        sys.exit(1)


if __name__ == "__main__":
    main()
