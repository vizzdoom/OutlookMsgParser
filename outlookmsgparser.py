import email.message
from convert_outlook_msg_file import outlookmsgfile
import uuid
from datetime import datetime
import os
import sys
import argparse
import shutil
import quopri


class OutlookMsgParser:
    def __init__(self, outlook_email_file_path, case_id=None):
        self.__log = ""
        self.__CASE_DIRECTORY_SUFFIX = "cases"
        self._original_email_file_location = os.path.abspath(outlook_email_file_path)
        try:
            self._print(f"[i] Outlook Email File Analysis for: {self._original_email_file_location}")
            self._emailMessage = outlookmsgfile.load(outlook_email_file_path)
        except FileNotFoundError:
            self.__log += f"[-] No such email file: {outlook_email_file_path}\r\n"
            sys.stderr(f"[-] No such email file: {outlook_email_file_path}")
            self.save_log()
            sys.exit(1)
        self.case_id = case_id or str(uuid.uuid4())
        self._case_directory = os.path.join(os.getcwd(), self.__CASE_DIRECTORY_SUFFIX, self.case_id)

        try:
            os.makedirs(self._case_directory, exist_ok=False)
        except OSError:
            self._print(f"[!] Case directory {self._case_directory} exists. Overwriting...")
            os.makedirs(self._case_directory, exist_ok=True)

        self._print(f"[i] Case ID: {self.case_id}")
        self._print(f"[i] Case directory: {self._case_directory}")
        self._print(f"[i] Time of analysis: {datetime.now()}")

    def _print(self, msg):
        self.__log += msg+"\r\n"
        print(msg+"\r\n")

    def copy_original_msg_to_case_directory(self):
        email_copy_path = os.path.join(self._case_directory, f"_orig__{os.path.basename(self._original_email_file_location)}")
        shutil.copyfile(self._original_email_file_location, email_copy_path)
        self._print(f"[i] Original email saved to: {email_copy_path}")

    def save_eml(self):
        email_path = os.path.join(self._case_directory, "email.eml")
        with open(email_path, mode="wb") as email_file:
            email_file.write(self._emailMessage.as_bytes())
        self._print(f"[i] Email in the eml format saved to: {email_path}")

    def save_payloads(self):
        payloads = self._emailMessage.get_payload()
        if isinstance(payloads, str):  # if we have only one body payload we want to case str to list with single item
            payloads = [payloads]
        self._print(f"[i] Detected payloads: {len(payloads)}")
        for payload_no, payload in enumerate(payloads):
            # do we have simple, string payload (not an attachment)?
            # if so, we do not have filename, let's save this just as .txt file
            # plus we can try to quote-printable decode this
            if isinstance(payload, str):
                payload_raw_filename = os.path.join(self._case_directory, f"payload_{payload_no}_rawstr.txt")
                payload_quoted_printable_filename = os.path.join(self._case_directory, f"payload_{payload_no}_decoded_quopri.txt")
                with open(payload_raw_filename, mode="w") as filehandler:
                    filehandler.write(payload)
                with open(payload_quoted_printable_filename, mode="wb") as filehandler:
                    filehandler.write(quopri.decodestring(payload))
            else:
                payload_attachment_filename = payload.get_filename() if payload.get_filename() is not None else "NOFILENAME"
                payload_raw_filename = os.path.join(self._case_directory, f"payload_{payload_no}_raw__{payload_attachment_filename}.txt")
                payload_decoded_filename = os.path.join(self._case_directory, f"payload_{payload_no}_decoded__{payload_attachment_filename}")
                if payload_attachment_filename == "NOFILENAME":
                    payload_decoded_filename += ".txt"

                with open(payload_raw_filename, encoding="utf-8", mode="w") as payload_raw_filehandler:
                    self._print(f"[i] Saving txt payload {payload_raw_filename}")
                    payload_raw_filehandler.write(payload.as_string())

                with open(payload_decoded_filename, encoding="utf-8", mode="w") as payload_decoded_filehandler:
                    self._print(f"[i] Saving txt payload {payload_raw_filename}")
                    for ph in payload._headers:
                        if ph == ('Content-Transfer-Encoding', 'quoted-printable'):
                            payload_decoded_filehandler.write(quopri.decodestring(payload.get_payload()).decode("utf-8"))
                            file_has_been_saved = True
                    if not file_has_been_saved:
                        payload_decoded_filehandler.write(payload.get_payload())

    def print_headers(self):
        headers_print_string = "[i] Printing email headers:\r\n"
        for h in self._emailMessage._headers:
            headers_print_string += f"{h[0]}:\t {h[1]}\r\n"
        self._print(headers_print_string)

    def save_log(self):
        with open(os.path.join(self._case_directory, "output.txt"), encoding="utf-8", mode="w") as file_log:
            file_log.write(self.__log)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Email Auditor parses email files in the .msg format and converts them to the .eml format. Thereafter, it saves headers and payloads into separate files for further, manual analysis.")
    parser.add_argument('msgfile', help="Path to the email in the .msg format")
    args = parser.parse_args()

    ea = OutlookMsgParser(outlook_email_file_path=args.msgfile)
    ea.copy_original_msg_to_case_directory()
    ea.save_eml()
    ea.save_payloads()
    ea.print_headers()
    ea._print(f"[i] Case {ea.case_id} analysis completed.")
    ea.save_log()
