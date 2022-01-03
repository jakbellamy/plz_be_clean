from pandas.io.clipboard import clipboard_get as read_clipboard
from appscript import app, k
from mactypes import Alias


def split_raw_email_list(list_string=read_clipboard()):
    return [email.split('>')[0] for email in list_string.split('<')[1:]]


class Outlook(object):
    def __init__(self):
        self.client = app("Microsoft Outlook")


class Message(object):
    def __init__(
            self,
            parent=None,
            subject="",
            body="",
            to_recipients=[],
            cc_recipients=[],
            bcc_recipients=[],
            show_=False,
    ):

        if parent is None:
            parent = Outlook()
        client = parent.client

        self.msg = client.make(
            new=k.outgoing_message,
            with_properties={k.subject: subject, k.content: body, k.sender: {k.name: 'Jakob Bellamy',
                                                                             k.address: 'Jakob.Bellamy@supremelending.com',
                                                                             k.type: k.unresolved_address}},
        )

        self.add_recipients(emails=to_recipients, type_="to")
        self.add_recipients(emails=cc_recipients, type_="cc")
        self.add_recipients(emails=bcc_recipients, type_="bcc")

        if show_:
            self.show()

    def show(self):
        self.msg.open()
        self.msg.activate()

    def add_attachment(self, p):
        # p is a Path() obj, could also pass string

        p = Alias(str(p))  # convert string/path obj to POSIX/mactypes path

        attach = self.msg.make(new=k.attachment, with_properties={k.file: p})

        return attach

    def add_recipients(self, emails, type_="to"):
        if not isinstance(emails, list):
            emails = [emails]
        for email in emails:
            self.add_recipient(email=email, type_=type_)

    def add_recipient(self, email, type_="to"):
        msg = self.msg

        if type_ == "to":
            recipient = k.to_recipient
        elif type_ == "cc":
            recipient = k.cc_recipient
        elif type_ == "bcc":
            recipient = k.bcc_recipient

        msg.make(new=recipient, with_properties={k.email_address: {k.address: email}})
