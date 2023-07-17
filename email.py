from O365 import Account

secret = '-l28Q~J-_qiBRL70h1ENJQ.REh43Ldedm60YNcJt'

client_id = 'd9a78337-a8dd-4cec-98ba-b2caf9489c81'

tenant_id = '9be73e3a-3b63-4ea5-8e18-675e687d2de9'

credentials = (client_id, secret)

scopes = ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send']


account = Account(credentials, auth_flow_type='credentials', tenant_id= tenant_id)
if account.authenticate():
   print('Authenticated!')

mailbox = account.mailbox(resource='hmartin11@telus.net')
message = mailbox.new_message()
message.to.add(['hmartin9@student.ubc.ca', 'martin.h.5@pg.com'])
#message.sender.address = 'hmartin11@telus.net'  # changing the from address
message.body = 'Hello World!!!!! This email was sent with Python'
#message.attachments.add('george_best_quotes.txt')
#message.save_draft()  # save the message on the cloud as a draft in the drafts folder

message.send() 

