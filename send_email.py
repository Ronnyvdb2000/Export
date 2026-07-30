"""
send_email.py

Stuurt alle Excel-bestanden (exports + validatierapporten) uit een map
als bijlage via SendGrid.

Benodigde environment variables (als GitHub secrets ingesteld):
  SENDGRID_API_KEY   - API-key van je SendGrid-account
  SENDGRID_FROM      - geverifieerd afzenderadres (single sender verification)
  NOTIFY_EMAIL       - ontvanger; komt normaal uit de workflow_dispatch input,
                       met een secret als fallback als er geen input is opgegeven
"""

import argparse
import base64
import glob
import os
import sys

from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition


def build_attachment(path: str) -> Attachment:
    with open(path, "rb") as f:
        data = f.read()
    encoded = base64.b64encode(data).decode()
    return Attachment(
        FileContent(encoded),
        FileName(os.path.basename(path)),
        FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        Disposition("attachment"),
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--attachments-dir", required=True,
                         help="map met alle .xlsx-bestanden die als bijlage moeten meegaan")
    parser.add_argument("--issues-found", type=int, default=0,
                         help="1 als een van de validaties issues vond, anders 0")
    args = parser.parse_args()

    api_key = os.environ.get("SENDGRID_API_KEY")
    sender = os.environ.get("SENDGRID_FROM")
    recipient = os.environ.get("NOTIFY_EMAIL")

    if not all([api_key, sender, recipient]):
        print("Fout: SENDGRID_API_KEY, SENDGRID_FROM en NOTIFY_EMAIL moeten allemaal ingesteld zijn.")
        sys.exit(1)

    files = sorted(glob.glob(os.path.join(args.attachments_dir, "*.xlsx")))
    if not files:
        print(f"Geen .xlsx-bestanden gevonden in {args.attachments_dir}")
        sys.exit(1)

    status = "⚠️ Validatieproblemen gevonden — zie rapport(en)" if args.issues_found else "✅ Validatie geslaagd, geen issues"
    bestandslijst = "".join(f"<li>{os.path.basename(f)}</li>" for f in files)

    message = Mail(
        from_email=sender,
        to_emails=recipient,
        subject=f"VLIF4-export klaar — {status}",
        html_content=(
            f"<p>De VLIF4-pipeline is uitgevoerd voor {len(files)} bestand(en).</p>"
            f"<p><b>Status:</b> {status}</p>"
            f"<ul>{bestandslijst}</ul>"
        ),
    )

    message.attachment = [build_attachment(f) for f in files]

    sg = SendGridAPIClient(api_key)
    response = sg.send(message)
    print(f"Mail verstuurd, statuscode: {response.status_code}")

    if response.status_code >= 300:
        sys.exit(1)


if __name__ == "__main__":
    main()
