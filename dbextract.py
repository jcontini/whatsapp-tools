import os, sys, time, sqlite3, openpyxl, datetime
from openpyxl import Workbook
from sys import argv
script, msgstore = argv

filters = {
	"subject": 'Xendit'
}

#Prep Excel File
def prep_excel():
	wb = Workbook()
	ws = wb.active
	#Create Chats Sheet
	ws.title = "Chat List"
	ws.append(['Chat','Datetime','key_remote_jid'])
	ws.freeze_panes = 'A2'
	column_widths = [30,15,35]
	for i, column_width in enumerate(column_widths):
		ws.column_dimensions[chr(65+i)].width = column_width
	#Create Messages Sheet
	ws = wb.create_sheet(index=0,title='Messages')
	ws.append(['Datetime','Chat','Sender','Message','Type','SenderID','Comments'])
	ws.freeze_panes = 'A2'
	column_widths = [15,20,20,50,20,20,50]
	for i, column_width in enumerate(column_widths):
		ws.column_dimensions[chr(65+i)].width = column_width

	return wb

#Go through each chatgroup
def loop_chats(wb):
	chats = [];i = 0
	print 'Scanning Chats...'
	c.execute("SELECT key_remote_jid, subject, sort_timestamp FROM chat_list")
	for line in c.fetchall():
		if filters['subject'] in str(line[1]):
			chats.append({
				'id':str(line[0]),
				'subject':str(line[1]),
				'timestamp':str(line[2]),
				})
			i+=1
			print "%s) %s" % (i,str(line[1]))

	ws = wb.get_sheet_by_name("Chat List")

	for chat in chats:
		print ('-'*40 + chat['subject'] + '-'*40)
		ws.append([chat['subject'],chat['timestamp'],chat['id']])
		extract_messages(chat['id'],chat['subject'])

#export every message in that chat to the excel format above
def extract_messages(chat_id,chat_subject):
	query = "SELECT key_remote_jid, key_from_me, data, remote_resource, received_timestamp FROM messages WHERE key_remote_jid='" + chat_id + "'"
	c.execute(query)
	i = 0
	ws = wb.get_sheet_by_name("Messages")
	for line in c.fetchall():
		i+=1
		if line[1] == 1:
			sender = "Me"
		else:
			try:
				sender = line[3].split('@')[0]
			except:
				sender = 'ERROR'
		try:
			message = unicode(line[2], 'utf8').encode('ascii', 'ignore')
		except:
			message = "ERROR"
		message_time = datetime.datetime.fromtimestamp(int(line[4])/1000)
		print "%s) %s" % (i, str(line[2]))
		ws.append([message_time.strftime('%Y-%m-%d %H:%M:%S'),chat_subject,sender,message,"",line[3],""])
	wb.save(xlsx_file)

#Connect to WhatsApp Database
conn = sqlite3.connect(msgstore)
c = conn.cursor()
conn.text_factory = str

xlsx_file = time.strftime("%Y-%m-%d") + '-WA.xlsx'
wb = prep_excel()

loop_chats(wb)
wb.save(xlsx_file)