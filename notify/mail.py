import getpass
import html
import io
import os
import re
import smtplib
import ssl
import zipfile
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from logs.logger_setup import get_execution_log

FILE_STATS = {}


def record_files(flow, found, done, failed):
    FILE_STATS[flow] = {
        "found": found,
        "done": done,
        "failed": failed
    }


def get_file_stats(flow):
    return FILE_STATS.get(flow)


class Notifier:
    """Trigger / exception / end mails. Never raises: a mail problem must not stop the bot."""

    def __init__(self, mail_cfg: dict):
        self.cfg = mail_cfg
        self.enabled = str(mail_cfg.get("mail_enabled", "")).strip().lower() in ("1", "true", "yes", "y")
        self.started = datetime.now()

    # ------------------------------------------------------------ helpers
    @staticmethod
    def _split(value):
        return [a.strip() for a in re.split(r"[;,]", value or "") if a.strip()]

    def _execution_attachment(self):
        """This run's execution file -> (filename, bytes, maintype, subtype), or None."""
        path = get_execution_log()
        if not path or not Path(path).exists():
            return None
        data = Path(path).read_bytes()
        if not data:
            return None
        if len(data) > 5 * 1024 * 1024:               # large log -> zip it
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr(Path(path).name, data)
            return Path(path).name + ".zip", buf.getvalue(), "application", "zip"
        return Path(path).name, data, "text", "plain"

    def _info(self):
        return f"Started  : {self.started:%Y-%m-%d %H:%M:%S}\n"

    def _send(self, kind, body, attach_log=True, html_body=None):
        if not self.enabled:
            return
        try:
            to = self._split(self.cfg.get("mail_to"))
            if not to:
                print("Mail skipped: mail_to is empty in config.xlsx")
                return
            password = self.cfg.get("sender_password", "").strip()
            if not password:
                print(f"Mail skipped: environment variable {self.cfg.get('mail_password_env')} is not set")
                return

            msg = EmailMessage()
            msg["Subject"] = f"{self.cfg['mail_subject_prefix']} {kind}"
            msg["From"] = self.cfg.get("mail_from") or self.cfg["smtp_user"]
            msg["To"] = ", ".join(to)
            cc = self._split(self.cfg.get("mail_cc"))
            if cc:
                msg["Cc"] = ", ".join(cc)

            msg.set_content(body)                                   # plain-text version
            if html_body:
                msg.add_alternative(html_body, subtype="html")      # HTML version

            if attach_log:
                att = self._execution_attachment()
                if att:
                    name, data, maintype, subtype = att
                    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)

            host, port = self.cfg["smtp_host"], int(self.cfg["smtp_port"])
            context = ssl.create_default_context()
            if port == 465:
                server = smtplib.SMTP_SSL(host, port, context=context, timeout=30)
            else:
                server = smtplib.SMTP(host, port, timeout=30)
                server.starttls(context=context)
            with server:
                server.login(self.cfg["smtp_user"], password)
                server.send_message(msg, to_addrs=to + cc)
            print(f"Mail sent: {kind}")
        except Exception as exc:
            print(f"Mail NOT sent ({kind}): {exc}")

    # ------------------------------------------------------------ HTML for the exception mail
    def _exception_html(self, flow, exc, traceback_text):
        e = html.escape
        rows = [
            ("Flow", flow),
            ("Started", f"{self.started:%Y-%m-%d %H:%M:%S}"),
            ("Failed at", f"{datetime.now():%Y-%m-%d %H:%M:%S}"),
        ]
        table_rows = "".join(
            f'<tr><td style="padding:6px 12px;background:#f3f4f6;font-weight:bold;width:110px;'
            f'border:1px solid #e5e7eb;">{e(label)}</td>'
            f'<td style="padding:6px 12px;border:1px solid #e5e7eb;">{e(str(value))}</td></tr>'
            for label, value in rows
        )
        return f"""\
<html>
<body style="margin:0;padding:16px;background:#f9fafb;font-family:Segoe UI,Arial,sans-serif;color:#111827;">
  <div style="max-width:760px;margin:auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;">
    <div style="background:#b91c1c;color:#ffffff;padding:14px 18px;font-size:18px;font-weight:bold;">
      &#9888; B&amp;W Bot &ndash; flow FAILED
    </div>
    <div style="padding:18px;">
      <p style="margin:0 0 12px 0;">The flow <b>{e(flow)}</b> stopped with an error.</p>
      <table style="border-collapse:collapse;font-size:14px;margin-bottom:16px;">{table_rows}</table>

      <div style="font-weight:bold;margin-bottom:6px;">Error</div>
      <div style="background:#fef2f2;border-left:4px solid #b91c1c;padding:10px 12px;margin-bottom:16px;font-size:14px;">
        {e(str(exc))}
      </div>

      <p style="margin:0;color:#6b7280;font-size:13px;"></p>
    </div>
  </div>
</body>
</html>"""

    # ------------------------------------------------------------ the three mails
    def trigger(self, flows):
        body = ("The B&W bot has started.\n\n" + self._info() +
                f"Flows    : {', '.join(flows)}\n")
        self._send("STARTED", body, attach_log=False)

    def exception(self, flow, exc, traceback_text):
        body = (f"The flow '{flow}' FAILED.\n\n" + self._info() +
                f"Error    : {exc}\n\n"
                f"Details  :\n{traceback_text}\n"
                "The execution log is attached.\n")
        self._send(f"FAILED - {flow}", body, html_body=self._exception_html(flow, exc, traceback_text))

    def end(self, results):
        failed_flows = [
            name for name, (status, _) in results.items()
            if status != "Success"
        ]

        took = datetime.now() - self.started

        lines = [
            f"  {name:<8} {status}" + (f"  -> {err}" if err else "")
            for name, (status, err) in results.items()
        ]

        # Collect file statistics
        stats_lines = []

        for flow in results:
            stats = get_file_stats(flow)

            if stats:
                stats_lines.append(
                    f"  {flow:<8} "
                    f"Found: {stats['found']} | "
                    f"Success: {stats['done']} | "
                    f"Failed: {stats['failed']}"
                )

        stats_section = (
            "\nFile Statistics:\n" + "\n".join(stats_lines) + "\n"
            if stats_lines
            else "\nFile Statistics:\nNo file statistics recorded.\n"
        )

        body = (
            f"The B&W bot has finished "
            f"{'WITH ERRORS' if failed_flows else 'successfully'}.\n\n"
            + self._info()
            + f"Duration : {str(took).split('.')[0]}\n\n"
            + "Result per flow:\n"
            + "\n".join(lines)
            + "\n"
            + stats_section
            + "\nThe execution log is attached.\n"
        )

        self._send(
            "COMPLETED - "
            + ("WITH ERRORS" if failed_flows else "SUCCESS"),
            body
        )
