
import PySimpleGUI as sg
from PySimpleGUI import filedialog
from pathlib import Path
import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
from win32com import client
import pandas as pd
import os
import time
from hashlib import md5
from PyPDF4 import PdfFileReader, PdfFileWriter
from PyPDF4.generic import NameObject, DictionaryObject, ArrayObject, \
    NumberObject, ByteStringObject
from PyPDF4.pdf import _alg33, _alg34, _alg35
from PyPDF4.utils import b_
from email.message import EmailMessage
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from password_generator import PasswordGenerator

# Pdf Encryption
def encrypt(writer_obj: PdfFileWriter, user_pwd, owner_pwd=None, use_128bit=True):
    """
    Encrypt this PDF file with the PDF Standard encryption handler.

    :param str user_pwd: The "user password", which allows for opening
        and reading the PDF file with the restrictions provided.
    :param str owner_pwd: The "owner password", which allows for
        opening the PDF files without any restrictions.  By default,
        the owner password is the same as the user password.
    :param bool use_128bit: flag as to whether to use 128bit
        encryption.  When false, 40bit encryption will be used.  By default,
        this flag is on.
    """
    import time, random
    if owner_pwd == None:
        owner_pwd = user_pwd
    if use_128bit:
        V = 2
        rev = 3
        keylen = int(128 / 8)
    else:
        V = 1
        rev = 2
        keylen = int(40 / 8)
    # permit copy and printing only:
    P = -44
    O = ByteStringObject(_alg33(owner_pwd, user_pwd, rev, keylen))
    ID_1 = ByteStringObject(md5(b_(repr(time.time()))).digest())
    ID_2 = ByteStringObject(md5(b_(repr(random.random()))).digest())
    writer_obj._ID = ArrayObject((ID_1, ID_2))
    if rev == 2:
        U, key = _alg34(user_pwd, O, P, ID_1)
    else:
        assert rev == 3
        U, key = _alg35(user_pwd, rev, keylen, O, P, ID_1, False)
    encrypt = DictionaryObject()
    encrypt[NameObject("/Filter")] = NameObject("/Standard")
    encrypt[NameObject("/V")] = NumberObject(V)
    if V == 2:
        encrypt[NameObject("/Length")] = NumberObject(keylen * 8)
    encrypt[NameObject("/R")] = NumberObject(rev)
    encrypt[NameObject("/O")] = ByteStringObject(O)
    encrypt[NameObject("/U")] = ByteStringObject(U)
    encrypt[NameObject("/P")] = NumberObject(P)
    writer_obj._encrypt = writer_obj._addObject(encrypt)
    writer_obj._encrypt_key = key


document_path = Path(__file__).parent / "Tem.docx"  # use the template here
doc = DocxTemplate(document_path)

date = datetime.datetime.today()

# set the theme for the screen/window
sg.theme('DarkBlue13')


def generate_from_excel():
    try:
        file_path = filedialog.askopenfilename()
        word_app = client.Dispatch("Word.Application")
        data_frame = pd.read_excel(file_path)
        for r_index, row in data_frame.iterrows():
            name = row['Name']
            # print(name)
            award_name = row['Award_Name']
            hours = row['Hours']
            location = row['Location']
            date = row['Date']
            email_address = row['Email']
            tpl = DocxTemplate("Tem.docx")
            df_to_dcot = data_frame.to_dict()
            x = data_frame.to_dict(orient='records')
            context = x
            tpl.render(context[r_index])
       
            '''if os.path.isfile(name + ".docx"):
                os.rename(name + ".docx", name + r_index + ".docx")
                tpl.save( name + r_index + ".docx")
            else:
                tpl.save(name + ".docx")'''
            r_index = str(r_index)
            tpl.save(name + r_index + ".docx")

           

            time.sleep(1)
            # get project folder path
            ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
            # convert docx to pdf
            doc = word_app.Documents.Open(ROOT_DIR  + '//' + name + r_index + '.docx')
            print('Exporting')
            doc.SaveAs(ROOT_DIR  + '//' +  name + r_index + '.pdf', FileFormat=17)

            unmeta = PdfFileReader(ROOT_DIR  + '//' + name + r_index + '.pdf')
            writer = PdfFileWriter()

            writer.appendPagesFromReader(unmeta)
           
            pwo = PasswordGenerator()
            #creates a randomly generated password for the student to acess the encrypted PDF
            user_pass = pwo.shuffle_password('qwertyuiopasdfghjklzxcvbnm1234567890', 5)
            #a set password for the creator of the PDF
            owner_pass = 'PSUabCE2022'

            encrypt(writer, user_pass, owner_pass)
            encrypted_pdf = ROOT_DIR  + '//' + name + r_index + "encrypted.pdf"
            os.remove(ROOT_DIR  + '//' +  name +  r_index + '.pdf')


            with open(encrypted_pdf, 'wb') as fp:
                writer.write(fp)
                #sg.popup("File saved", f"File has been saved here: {encrypted_pdf}")

            # Email Code
            msg = MIMEMultipart()
            sender = "testtest1205@outlook.com"
            password = "abc1234567890"
            receiver = email_address
            '''body = f"Congratulations {name} on your {award_name}!!" \
                    f"\nYour official Penn State Certificate is attached\nYour password for your certificate is:  " + user_pass'''
            body =  f'<pre>On behalf of the <a href="https://www.abington.psu.edu/continuing-education">Penn State Abington Continuing Education</a> team,\
congratulations on the completion of the {award_name} program. \
You will find your certificate of program completion attached to this email. \
The password to unlock your certificate is: <u>{user_pass}</u>. If you have any questions, please contact us at abce@psu.edu. \
Stay connected with us on <a href="https://www.linkedin.com/company/penn-state-abington-continuing-education/?viewAsMember=true ">LinkedIn</a> \
and <a href = "https://www.facebook.com/PSUAbingtonCE/">Facebook<a/>.<pre>'
            msg.attach(MIMEText(body, 'html'))
            msg['Subject'] = f'{award_name} Certificate Completion'
            msg['From'] = sender
            msg['To'] = receiver
            binary_pdf = open(encrypted_pdf, 'rb')
       
            payload = MIMEBase('application', 'octate-stream', Name = name + r_index + '.pdf')
            payload.set_payload(binary_pdf.read())
       
            # enconding the binary into base64
            encoders.encode_base64(payload)
       
            # add header with pdf name
            payload.add_header('Content-Decomposition', 'attachment', filename = name + r_index + '.pdf')
            msg.attach(payload)
       
            with smtplib.SMTP('smtp.office365.com', 587) as server:
                server.starttls()
                server.login(sender, password)
                server.sendmail(sender, receiver, msg.as_string())
                #Pop up verifying email was sent
                #sg.popup("Certificate Has been emailed to: " + email_address)

        sg.popup("All Certificates have been emailed")

        word_app.Quit()
    except Exception as e:
        sg.Popup("Excel format Incorrect")


def open_window():
    layout = [
        [sg.Text("Name"), sg.Input(size=(45, 20), key="Name", do_not_clear=False)],
        [sg.Text("Award Name"),
         sg.Combo(['Digital Marketing Professional','Diversity, Equity, Inclusion, and Belonging in the Workplace','Nursing Care Home Administrator','Personal Care Home Administrator','Project Management','Public Entity Leadership Development','Trauma Informed Practices for Educators'], size=(43, 20), key="Award_Name")],
        [sg.Text("Hours Completed"), 
         sg.Combo(['2','7','15','30','49','84','100','120'], size=(43, 20), key="Hours", )],
        [sg.Text("Location"), sg.Combo(['Virtual', 'Delaware Valley Trusts'], size=(43, 20), key="Location")],
        [sg.Text("email"), sg.Input(key="Email", size=(45, 20), do_not_clear=False)],
        [sg.Checkbox("Email Certificate Now", default=False, key="Emailed Certificate")],
        [sg.Button("Generate Certificate"), sg.Exit("Go Back")],
    ]
    window = sg.Window('Certificate Details', layout, modal= True, element_justification="right", resizable = True)
    choice = None
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Go Back":
            break
        if event == "Generate Certificate":
            print(event, values)
            values["Date"] = date.strftime("%B %Y")
            doc.render(values)
            output_path = Path(__file__).parent / f"{values['Name']}.docx"
            doc.save(output_path)
            pdf_output_path = Path(__file__).parent / f"{values['Name']}.pdf"
            convert(output_path, pdf_output_path)
            sg.popup("File saved", f"File has been saved here: {pdf_output_path}")
            time.sleep(1)
            # get project folder path
            ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
            unmeta = PdfFileReader(ROOT_DIR + f"//{values['Name']}.pdf")
            writer = PdfFileWriter()

            writer.appendPagesFromReader(unmeta)
           
            pwo = PasswordGenerator()
            #creates a randomly generated password for the student to acess the encrypted PDF
            user_pass = pwo.shuffle_password('qwertyuiopasdfghjklzxcvbnm1234567890', 5)
            #a set password for the creator of the PDF
            owner_pass = 'PSUabCE2022'

            encrypt(writer, user_pass, owner_pass)
            encrypted_pdf = ROOT_DIR + f"//{values['Name']}encrypted.pdf"
            os.remove(ROOT_DIR  + '//' + values['Name'] + '.pdf')

            with open(encrypted_pdf, 'wb') as fp:
                writer.write(fp)
            sg.popup("File saved", f"File has been saved here: {encrypted_pdf}")
            # Email Code
            if values["Emailed Certificate"] == True:
                    msg = MIMEMultipart()
                    sender = "testtest1205@outlook.com"
                    password = "abc1234567890"
                    receiver = values['Email']
                    '''body = f"Congratulations {values['Name']} on your {values['Award_Name']} with a total of {values['Hours']} hours at {values['Location']}!!" \
                           f"\nYour official certificate is attached" + "\nYour password for your certificate is:  " + user_pass '''
                    body = f'<pre>On behalf of the <a href="https://www.abington.psu.edu/continuing-education">Penn State Abington Continuing Education</a> team,\
congratulations on the completion of the {values["Award_Name"]} program. \
You will find your certificate of program completion attached to this email. \
The password to unlock your certificate is: <u>{user_pass}</u>. If you have any questions, please contact us at abce@psu.edu. \
Stay connected with us on <a href="https://www.linkedin.com/company/penn-state-abington-continuing-education/?viewAsMember=true ">LinkedIn</a> \
and <a href = "https://www.facebook.com/PSUAbingtonCE/">Facebook<a/>.<pre>'
                    msg.attach(MIMEText(body, 'html'))
                    msg['Subject'] = f'Certificate for completion'
                    msg['From'] = sender
                    msg['To'] = receiver
                    binary_pdf = open(encrypted_pdf, 'rb')
               
                    payload = MIMEBase('application', 'octate-stream', Name=f"{values['Name']}.pdf")
                    payload.set_payload(binary_pdf.read())
               
                    # enconding the binary into base64
                    encoders.encode_base64(payload)
               
                    # add header with pdf name
                    payload.add_header('Content-Decomposition', 'attachment', filename=f"{values['Name']}.pdf")
                    msg.attach(payload)
               
                    with smtplib.SMTP('smtp.office365.com', 587) as server:
                        server.starttls()
                        server.login(sender, password)
                        server.sendmail(sender, receiver, msg.as_string())
                        #Pop up verifying email was sent
                        sg.popup("Certificate Has been emailed to: " + values['Email'])
    window.close()

# define layout
selection = [[sg.Text('Select an option to Create Certificates', size=(31, 1), font='Lucida', justification='left')],
          [sg.Combo(['Using Excel Sheet', 'Entering Details'], enable_events=True,
                     default_value='Using Excel Sheet', key='choice')],
          [sg.Button('SELECT', font=('Times New Roman', 12)), sg.Button('CANCEL', font=('Times New Roman', 12))]]

window = sg.Window('Certificate Generator', selection, resizable = True)

while True:
    event, values = window.Read()
    if event is None or event == 'CANCEL':
        break

    if event == 'SELECT':
        combo = values['choice']  # use the combo key
        if combo == 'Using Excel Sheet':
            generate_from_excel()
            # print(combo)
        elif combo == 'Entering Details':
            open_window()

            # print(combo)


window.Close()
