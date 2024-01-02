import email
import imaplib
import os
import re
import tkinter as tk
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from tkinter import ttk

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
from tkcalendar import Calendar


def get_selected_date():
    global date  # Declare 'date' as a global variable
    selected_date_str = cal.get_date()
    selected_date = datetime.strptime(selected_date_str, "%x")
    date = selected_date.strftime("%d.%m.%Y")
    print("Selected Date:", date)
    root.destroy()


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


user = input("Pls enter you User ID : ")  # "trade@smartserve.co"
password = input("Pls enter your Password : ")  # "Trade@1234"
date = date  # input("Pls select date like (06-oct-2023) : ") # "06-oct-2023"


imap_url = "smartserve.co"  #'outlook.office365.com' #smartserve.co

my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(user, password)

print("imap.list()", my_mail.list())
sent = {}
count = 0
for Folder in ["Inbox.Sent", "INBOX"]:
    my_mail.select(Folder)

    # Construct the search query using the HEADER criterion
    search_query = f"(SUBJECT {date})"

    _, search = my_mail.search(None, search_query)
    # _, search = my_mail.search(None, f'(SENTON {date})')

    msgs = []
    for msg in search[0].split():
        _, data = my_mail.fetch(msg, "(RFC822)")
        msgs.append(data)

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

                cont = my_msg["subject"].replace("Re", "")

                cont = re.sub("[^a-zA-Z0-9.]", "", cont)
                # cont=cont.replace('Please Approve the Trade', '')
                cont = cont.split("Please", 1)[0]
                extracted_word = extracted_word.replace("+canned_response", "")
                print(extracted_word)
                doc_folder = "doc"
                if not os.path.exists(doc_folder):
                    os.makedirs(doc_folder)

                if Folder == "Inbox.Sent":
                    sent[f"{extracted_word}_{cont}".upper()] = my_msg

                else:
                    pdf_filename = os.path.join("Doc", f"{extracted_word}_{cont}.pdf")

                    c = canvas.Canvas(pdf_filename, pagesize=letter)

                    # Add text to PDF
                    c.drawString(
                        20,
                        720,
                        f'From:       {my_msg["from"].split(" ")[0]}" <{extracted_word.replace("_", ".")}>',
                    )
                    input_timestamp = my_msg.get("Received")
                    # last_received_entry = input_timestamp[-1]
                    input_timestamp = input_timestamp.split(";")[-1].strip()
                    input_timestamp = parsedate_to_datetime(input_timestamp)

                    print(input_timestamp, "input_timestamp")
                    input_datetime = datetime.strptime(
                        str(input_timestamp), "%Y-%m-%d %H:%M:%S%z"
                    )
                    target_timezone = timezone(timedelta(hours=5, minutes=30))
                    converted_datetime = input_datetime.astimezone(target_timezone)
                    result_timestamp = converted_datetime.strftime("%d %b %Y %H:%M")
                    c.drawString(20, 700, "Date:        {}".format(result_timestamp))
                    c.drawString(20, 680, "To:           {}".format(my_msg["To"]))

                    c.drawString(20, 660, "Subject:   {}".format(my_msg["subject"]))

                    y_position = 600
                    for part in my_msg.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True)
                            body_text = body.decode("utf-8", "ignore")
                            for line in body_text.split("\n"):
                                if "trade@smartserve.co" not in line:
                                    line = line.replace(">", "", 1)
                                c.drawString(20, y_position, line[:-1])
                                result = re.search("[a-zA-Z0-9]", line)
                                if result:
                                    y_position -= 18
                                else:
                                    y_position -= 12

                    y_position -= 40

                    input_timestamp = sent[f"{extracted_word}_{cont}".upper()].get(
                        "Date"
                    )
                    input_datetime = datetime.strptime(
                        input_timestamp, "%a, %d %b %Y %H:%M:%S %z"
                    )
                    target_timezone = timezone(timedelta(hours=5, minutes=30))
                    converted_datetime = input_datetime.astimezone(target_timezone)
                    result_timestamp = converted_datetime.strftime(
                        "%a, %d %b %Y at  %H:%M"
                    )

                    c.drawString(
                        20,
                        y_position,
                        f"on {result_timestamp} {sent[f'{extracted_word}_{cont}'.upper()]['from']} Wrote: ",
                    )
                    starting = y_position
                    y_position -= 50
                    for part in sent[f"{extracted_word}_{cont}".upper()].walk():
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
                    c.setLineWidth(0.2)
                    c.line(15, starting - 10, 15, y_position)

                    c.save()

        print(count)
        print(len(sent))

    my_mail.close()
