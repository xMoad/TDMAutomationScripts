import win32com.client

#Connect to application
ol = win32com.client.Dispatch('Outlook.Application')

# size of the new email
olmailitem = 0x0
 
 
newmail = ol.CreateItem(olmailitem)

newmail.Subject = 'Testing Mail'

newmail.To = 'xyz@example.com'
newmail.CC = 'xyz@example.com'

newmail.Body= 'Hello, this is a test email.'

# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)

# To display the mail before sending it
# newmail.Display() 

newmail.Send()