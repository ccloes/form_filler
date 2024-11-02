import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
mail = outlook.CreateItem(0)
mail.To = "ccloes@gmail.com"
mail.Subject = "Automated Email from Python"
mail.Body = "This is an automated email created using Python."
mail.Save() # Save email as a draft instead of sending it
