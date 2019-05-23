
def Emailer(text, subject, recipient):
    import win32com.client as win32

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    # mail.send

Emailer("text","subject h","habib@gmail.com")

