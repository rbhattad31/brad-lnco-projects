# import pywintypes               # this package is installed by installing pywin32
# from win32com import client     # this package is installed by installing pywin32
#
#
# def send_mail(to, cc, subject, body):
#     try:
#         outlook = client.Dispatch('outlook.application')
#         mail = outlook.CreateItem(0)
#         mail.To = to
#         mail.cc = cc
#         mail.Subject = subject
#         mail.Body = body
#         mail.Send()
#     except pywintypes.com_error as message_error:
#         print("Sendmail error - Please check outlook connection")
#         return message_error
#     except Exception as error:
#         return error


def send_mail(to, cc, subject, body):
    # testing purpose only
    # delete this function, and uncomment above code
    pass
