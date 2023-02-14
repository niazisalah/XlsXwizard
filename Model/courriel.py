import outlook


o = outlook.Outlook()


email = o.inbox().get_emails()[0]


for attachment in email.attachments:
    attachment.download()

