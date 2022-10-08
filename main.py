import win32com.client as win32
from bs4 import BeautifulSoup
import pandas as pd
import os
import re

outlook = win32.Dispatch('Outlook.Application')

# Save your word document as a FILTERED htm file, and put the name here
save_name = "mail2"

html_file = f"{save_name}.htm"
attachments_folder = f"{save_name}_files"

# Open mail.xlsx in pandas; This contains columns to fill in the template
df = pd.read_excel('data.xlsx')

# Loop through the rows of the dataframe
for index, row in df.iterrows():
    # Send mail.html (which makes use of images in mail_files) to the email address in the dataframe
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem
    mail = outlook.CreateItem(0)

    mail.To = row['Mail']
    # mail.CC = row['CC']
    # mail.BCC = row['BCC']

    mail.From = 'arno.deceuninck@student.uantwerpen.be'
    mail.Subject = 'Subject'  # DON'T FORGET TO CHANGE THIS

    # Open the html file and read it
    html = open(html_file, 'r').read()

    # find regexes of the form «*» in the html file and replace them with the corresponding value in the dataframe
    for match in re.findall(r'«(.+?)»', html):
        assert match in row.index, f"Column {match} not found in dataframe"
        html = html.replace(f'«{match}»', str(row[match]))

    # Change the location of the images, since they become attachments
    html = html.replace(f"src=\"{attachments_folder}/", f"src=\"")

    mail.HTMLBody = html
    # Convert the html to text (in case the mail client doesn't support html)
    text = BeautifulSoup(html, 'html.parser').get_text()
    # Remove the empty lines at the start and end of the text
    text = text.strip()
    # mail.Body = text # Don't use this, it will be the default!

    # add the images from mail_files that are used in mail.htm
    for file in os.listdir(attachments_folder):
        abspath = os.path.abspath(os.path.join(attachments_folder, file))
        mail.Attachments.Add(abspath)

    mail.Send()
