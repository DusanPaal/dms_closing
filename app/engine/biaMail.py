# pylint: disable = C0103, R1711, W1203

"""
The 'biaMail.py' module
    - creates and sends of emails via SMTP server without
      the need of using an installed local email client.
    - fetches a Message object from a mailbox by connecting
      to MS Exchange

Version history:
1.0.20220504 - removed unused virtual key mapping fom _vkeys {}
1.0.20220906 - Minor code style improvements. The code section
               responsible for attaching files to message moved
               from send_message() to a separate _attach_file()
               procedure.
"""

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
from os.path import isfile, split
import re
import socket
from smtplib import SMTP
from typing import Union

# custom message classes
class SmtpMessage(MIMEMultipart):
    """
    A wrapper for MIMEMultipart
    objects that are sent using
    SMTP server.
    """

# custom warnings
class UndeliveredWarning(Warning):
    """
    Raised when message is not
    delivered to all recipients.
    """

# custom exceptions
class AttachmentNotFoundError(Exception):
    """
    Raised when a file attachment
    is requested but doesn't exist.
    """

class AttachmentSavingError(Exception):
    """
    Raised when any exception is
    caught durng writing of attachment
    data to a file.
    """

class InvalidSmtpHostError(Exception):
    """
    Raised when an invalid host name
    is used for SMTP connection.
    """

_logger = logging.getLogger("master")

def _sanitize_emails(addr: Union[str,list]) -> list:
    """
    Trims email addresses and validates
    the correctness of their email format
    according to the company's standard.
    """

    if not isinstance(addr, str) and len(addr) == 0:
        raise ValueError("No message recipients provided in 'to_addr' argument!")

    mails = []
    validated = []

    if isinstance(addr, str):
        mails = [addr]
    elif isinstance(addr, list):
        mails = addr
    else:
        raise TypeError(f"Argument 'addr' has invalid type: {type(addr)}")

    for mail in mails:

        stripped = mail.strip()
        validated.append(stripped)

        # check if email is Ledvance-specific
        match = re.search(r"\w+\.\w+@ledvance.com", stripped)

        if match is not None:
            continue

        _logger.warning(f"Possibly invalid email address used: '{stripped}'.")

    return validated

def _attach_files(email: SmtpMessage, att_paths: list) -> SmtpMessage:
    """
    Attaches files to a SmtpMessage object.
    """

    for att_path in att_paths:

        # Check whether email
        # attachmment path exists
        if not isfile(att_path):
            raise AttachmentNotFoundError(f"Attachment not found: {att_path}")

        with open(att_path, "rb") as file:
            payload = file.read()

        # The content type "application/octet-stream" means
        # that a MIME attachment is a binary file
        part = MIMEBase("application", "octet-stream")
        part.set_payload(payload)
        encoders.encode_base64(part)

        # get file name
        file_name = split(att_path)[1]

        # Add header
        part.add_header(
            "Content-Disposition",
            f"attachment; filename = {file_name}"
        )

        # Add attachment to the message
        # and convert it to a string
        email.attach(part)

    return email

def create_message(from_addr: str, to_addr: Union[str,list], subj: str,
                   body: str, att: Union[str,list] = None) -> SmtpMessage:
    """
    Creates a message with or without attachment(s).

    Params:
    -------
    from_addr:
        Email address of the sender.

    to_addr:
        Email address(es) of recipient(s). \n
        If a single email address
        is used, the message will be sent to that specific address. \n
        If multiple addresses are used, then the message will be sent
        to all of the recipients.

    subj:
        Message subject.

    body:
        Message body in HTML format.

    att:
        Any valid path(s) to message atachment file(s). \n
        If None is used (default), then message will be created without any file attached. \n
        If a file path is passed, then this file will be attached to the message. \n
        If multiple paths are used, these will be attached as multiple attachments to the message.

    Returns:
    --------
    A SmtpMessage object representing the message.

    Raises:
    -------
    AttachmentNotFoundError:
        If any of the attachment paths used is not found.
    """

    # sanitize input
    recips = _sanitize_emails(to_addr)

    # process
    email = MIMEMultipart()
    email["Subject"] = subj
    email["From"] = from_addr
    email["To"] = ";".join(recips)
    email.attach(MIMEText(body, "html"))

    if att is None:
        return email

    if isinstance(att, list):
        att_paths = att
    elif isinstance(att, str):
        att_paths = [att]
    else:
        raise TypeError(f"Argument 'att' has invalid type: {type(att)}")

    email = _attach_files(email, att_paths)

    return email

def send_smtp_message(msg: SmtpMessage, host: str, port: int):
    """
    Sends a message using SMTP server.

    Params:
    -------
    msg:
        An exchengelib Message object representing the Message to be sent.

    host:
        Name of the SMTP host server used for message sending.

    port:
        Number o the SMTP server port.

    Returns:
    --------
    None.

    Raises:
    -------
    UndeliveredWarning:
        When message fails to reach all the required recipients.

    TimeoutError:
        When attempting to connect to the SMTP server times out.
    """

    try:
        with SMTP(host, port, timeout = 30) as smtp_conn:
            smtp_conn.set_debuglevel(0) # off = 0; verbose = 1; timestamped = 2
            send_errs = smtp_conn.sendmail(msg["From"], msg["To"].split(";"), msg.as_string())
    except socket.gaierror as exc:
        raise InvalidSmtpHostError(f"Invalid SMTP host name: {host}") from exc
    except TimeoutError as exc:
        raise TimeoutError("Attempt to connect to the SMTP servr timed out! Possible reasons: "
        "Slow internet connection or an incorrect port number used.") from exc

    if len(send_errs) != 0:
        failed_recips = ';'.join(send_errs.keys())
        raise UndeliveredWarning(f"Message undelivered to: {failed_recips}")

    return
