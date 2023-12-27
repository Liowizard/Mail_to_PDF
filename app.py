import email
import imaplib
import os
import re
import tkinter as tk
from datetime import date, datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
from tkcalendar import Calendar

# from tkinter import messagebox


def get_selected_date():
    global date  # Declare 'date' as a global variable
    selected_date_str = cal.get_date()
    selected_date = datetime.strptime(selected_date_str, "%x")
    date = selected_date.strftime("%d.%m.%Y")
    print("Selected Date:", date)
    root.destroy()


def get_user_input(choice):
    global a  # Access the 'a' variable declared outside the function

    if choice == "Yes":
        a = "INBOX"
    elif choice == "No":
        a = "Inbox.Sent"

    root.destroy()


root = tk.Tk()
root.title("Yes or No Dialog")
root.geometry("300x150")

label = tk.Label(root, text="Choose Your required folder :")
label.pack(pady=10)

yes_button = tk.Button(root, text="Inbox", command=lambda: get_user_input("Yes"))
yes_button.pack(pady=10)

no_button = tk.Button(root, text="Sent", command=lambda: get_user_input("No"))
no_button.pack(pady=10)

root.mainloop()


root = tk.Tk()
root.title("Date Picker")


current_date = date.today()
cal = Calendar(
    root,
    selectmode="day",
    year=current_date.year,
    month=current_date.month,
    day=current_date.day,
)
cal.pack(pady=10)

select_button = tk.Button(root, text="Select Date", command=get_selected_date)
select_button.pack(pady=10)

date = None  # Initialize the 'date' variable

root.mainloop()


user = "trade@smartserve.co"  # input("Pls enter you User ID : ")  # "ajeshrandam@gmail.com"  AjeshTest6@outlook.com   trade@smartserve.co
password = "Trade@1234"  # input("Pls enter your Password : ")  # "gcic sssf zsfm bvox" Ajesh1161      Trade@1234
date = date  # input("Pls select date like (06-oct-2023) : ") # "06-oct-2023"
Folder = a  # input("Inbox or Sent : ")

imap_url = "smartserve.co"  #'outlook.office365.com' #smartserve.co

my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(user, password)

print("imap.list()", my_mail.list())

my_mail.select(Folder)

# Construct the search query using the HEADER criterion
search_query = f"(SUBJECT {date})"

_, search = my_mail.search(None, search_query)
# _, search = my_mail.search(None, f'(SENTON {date})')

msgs = []
for msg in search[0].split():
    _, data = my_mail.fetch(msg, "(RFC822)")
    msgs.append(data)


count = 0
for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple:
            my_msg = email.message_from_bytes((response_part[1]))
            print("my_msg", my_msg["to"])
            if Folder == "Inbox.Sent":
                name = my_msg["to"]
            elif Folder == "INBOX":
                name = my_msg["from"]
            match = re.search(r"<(.*?)>", name)
            if match:
                extracted_word = match.group(1).replace(".", "_")
            else:
                extracted_word = name.replace(".", "_")
            cont = my_msg["subject"].replace("Re: Trade confirmation ", "")
            cont = cont.replace("Re: Trade Confirmation ", "")
            cont = re.sub("[^a-zA-Z0-9.]", "", cont)
            # cont=cont.replace('Please Approve the Trade', '')
            cont = cont.split("Please", 1)[0]
            print(cont)
            extracted_word = extracted_word.replace("+canned_response", "")
            print(extracted_word)
            count += 1

            doc_folder = "doc"
            if not os.path.exists(doc_folder):
                os.makedirs(doc_folder)
            pdf_filename = os.path.join("Doc", f"{extracted_word}_{cont}_{count}.pdf")

            c = canvas.Canvas(pdf_filename, pagesize=letter)

            # Add text to PDF
            c.drawString(20, 720, "Subject: {}".format(my_msg["subject"]))
            c.drawString(
                20,
                700,
                f'From: {my_msg["from"].split(" ")[0]}" <{extracted_word.replace("_", ".")}>',
            )
            c.drawString(20, 680, "To: {}".format(my_msg["To"]))
            date = my_msg["Date"].split("-")[0]
            c.drawString(20, 660, "Date: {}".format(date))
            y_position = 620
            for part in my_msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True)
                    body_text = body.decode("utf-8", "ignore")
                    for line in body_text.split("\n"):
                        if "trade@smartserve.co" not in line:
                            line = line.replace(">", "", 1)
                        # print(line)
                        c.drawString(20, y_position, line[:-1])
                        result = re.search("[a-zA-Z0-9]", line)
                        if result:
                            y_position -= 12
                        else:
                            y_position -= 6
            print("-----------------------------------------------")

            c.setFont("Helvetica-Bold", 12)
            c.drawString(20, 750, "trade@smartserve.co")
            # Add lines to PDF
            c.setLineWidth(3)
            c.setStrokeColor(colors.black)
            c.line(20, 745, 550, 740)
            c.setLineWidth(0.5)
            c.line(20, 655, 550, 655)

            c.save()

print(count)

my_mail.close()
