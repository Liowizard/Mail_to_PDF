import imaplib
import email
import os
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import Image
import tkinter as tk
from tkcalendar import Calendar
from datetime import datetime



def get_selected_date():
    global date  # Declare 'date' as a global variable
    selected_date_str = cal.get_date()
    selected_date = datetime.strptime(selected_date_str, "%x")
    date = selected_date.strftime("%d.%m")
    print("Selected Date:", date)
    root.destroy()




root = tk.Tk()
root.title("Date Picker")

cal = Calendar(root, selectmode="day", year=2023, month=10, day=6)
cal.pack(padx=10, pady=10)

select_button = tk.Button(root, text="Select Date", command=get_selected_date)
select_button.pack(pady=10)

date = None  # Initialize the 'date' variable

root.mainloop()



 
user = "trade@smartserve.co"#input("Pls enter you User ID : ")  # "ajeshrandam@gmail.com"  AjeshTest6@outlook.com   trade@smartserve.co
password ="Trade@1234" #input("Pls enter your Password : ")  # "gcic sssf zsfm bvox" Ajesh1161      Trade@1234
date =date #input("Pls select date like (06-oct-2023) : ") # "06-oct-2023"
Folder="Inbox"# input("Inbox or Sent : ")

imap_url = "smartserve.co"#'outlook.office365.com' #smartserve.co

my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(user, password)

print("imap.list()",my_mail.list())

my_mail.select(Folder)

# Construct the search query using the HEADER criterion
search_query = f'(SUBJECT {date})'

_, search = my_mail.search(None, search_query)
# _, search = my_mail.search(None, f'(SENTON {date})')

msgs = []
for msg in search[0].split():
    _, data = my_mail.fetch(msg, "(RFC822)")
    msgs.append(data)



count=[]
for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg = email.message_from_bytes((response_part[1]))
            match = re.search(r'<(.*?)>', my_msg["from"])
            if match:
                extracted_word = match.group(1).replace('.', '_')
            else:
                extracted_word = my_msg["from"].replace('.', '_')
            cont=my_msg["subject"].replace('Re: Trade confirmation ', '')
            cont=cont.replace('Re: Trade Confirmation ', '')
            cont=re.sub("[^a-zA-Z0-9.]", "", cont)
            # cont=cont.replace('Please Approve the Trade', '')
            cont=cont.split("Please", 1)[0]
            print(cont)
            extracted_word=extracted_word.replace('+canned_response','')
            print(extracted_word)
            count.append(extracted_word)

            doc_folder = 'doc'
            if not os.path.exists(doc_folder):
                os.makedirs(doc_folder)
            pdf_filename =os.path.join("Doc", f'{extracted_word}_{cont}.pdf')

            c = canvas.Canvas(pdf_filename, pagesize=letter)

            # Add text to PDF
            c.drawString(50, 715, "Subject: {}".format(my_msg["subject"]))
            c.drawString(50, 695, f'From: {my_msg["from"].split(" ")[0]}" <{extracted_word.replace("_", ".")}>' )
            y_position = 640
            for part in my_msg.walk():
                if part.get_content_type() == 'text/plain':
                    body = part.get_payload(decode=True)
                    body_text = body.decode('utf-8', 'ignore')
                    for line in body_text.split('\n'):
                        if "trade@smartserve.co" not in line: 
                            line=line.replace('>', '',1)
                        # print(line)
                        c.drawString(50, y_position, line[:-1])
                        result = re.search("[a-zA-Z0-9]", line)
                        if result:
                            y_position -= 12
                        else:
                             y_position -= 6
            print("-----------------------------------------------")


            c.setFont("Helvetica-Bold", 12)
            c.drawString(50, 750, "trade@smartserve.co")
            # Add lines to PDF
            c.setLineWidth(3)
            c.setStrokeColor(colors.black)
            c.line(50, 745, 550, 740)
            c.setLineWidth(0.5)
            c.line(50, 655, 550,655)

            c.save()

print(len(count))

my_mail.close()



