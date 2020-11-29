# pandas is for read the files
import pandas as pd
import numpy as np
# easygui is for the gui
from easygui import *
# for pdf
from fpdf import FPDF
# packages for send the mail
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path
import mycredits


class CompareModifiyData():

    def take_input(self):
        """take files from the user from gui elemnents"""
        start_code = buttonbox("LET'S start the code,by selecting files",
                               "send_pdf", ("select files",))

        if start_code:
            try:
                # reading a file
                file1 = fileopenbox()
                file2 = fileopenbox()
                # have to read any csv or excel file
                file1_type, file2_type = \
                    file1.split('.')[-1], file2.split('.')[-1]
                df1, df2 = None, None
                _templist = [(file1_type, [df1], file1),
                             (file2_type, [df2], file2)]

                for value in _templist:
                    if value[0] == "csv":
                        value[1][0] = pd.read_csv(value[2])

                    elif value[0] in ["xlsx", "xls"]:
                        value[1][0] = pd.read_excel(value[2])
                    else:
                        raise Exception

            except Exception as e:
                exceptionbox(
                    msg="provide .csv or .xls or .xlsx format files only.",
                    title="provide valid file type")
        else:
            # exit the code
            print("exit the program")
        return _templist[0][1][0], _templist[1][1][0]

    def read_files(self):
        df1, df2 = self.take_input()
        try:
            # when the rows are same only data is change
            ne_stacked = (df1 != df2).stack()
            changed = ne_stacked[ne_stacked]
            # inittialize the column name to display
            changed.index.names = ['Row Number', 'Column name']
            # getting what values have changes
            difference_locations = np.where(df1 != df2)
            changed_from = df1.values[difference_locations]
            change_to = df2.values[difference_locations]
            # creating the Dataframe with filtered data
            var = pd.DataFrame(
                {'from': changed_from, 'to': change_to}, index=changed.index)
            df_type = "columns"
        except Exception as e:
            # to get latest added rows
            var = pd.concat([df1, df2]).drop_duplicates(keep=False)
            # to know the what type of data sending
            df_type = "rows"
        return var, df_type

# creating pdf_style class


class pdf_style(FPDF):
    def header(self):
        """setting header styles"""
        self.set_font('Arial', 'I', 12)  # setting a font
        self.cell(80)  # create a cell
        # (cell width, font, text, style,cell_positon,'text positon')
        self.cell(30, 10, 'Data', 1, 0, 'C')
        self.ln(20)  # line break

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, 'Page' + str(self.page_no()) + '/{nb}', 0, 0, 'C')


# pdf_creation
def create_pdf(data=None, data_type=None, file_path='nochange.pdf'):
    pdf = pdf_style(format='A4')
    pdf.alias_nb_pages()
    pdf.add_page()
    page_width = pdf.w - 2 * pdf.l_margin
    col_width = page_width / 4
    head_size, col_size = 12, 10
    pdf.set_font('Times', '', head_size)
    if data is not None:
        if data_type == 'columns':
            pdf.cell(50, head_size,
                     "Data change between common columns", 0, 0, 'C')
            pdf.ln(head_size)
            # formating the column names
            for col_value in ['Row Numbers', 'Column Names', 'From', 'To']:
                pdf.cell(col_width, col_size, col_value, 1, 0, 'C')
            pdf.ln(col_size)
            # formating the data
            for i in range(len(data.index)):
                pdf.cell(col_width, col_size,
                         str(data.index.get_level_values(
                             'Row Number').values[0]),
                         1, 0, 'C')
                pdf.cell(col_width, col_size, str(
                    data.index.get_level_values('Column name').values[0]),
                     1, 0, 'C')
                pdf.cell(col_width, col_size, str(
                    data['from'].iloc[i]), 1, 0, 'C')
                pdf.cell(col_width, col_size, str(
                    data['to'].iloc[i]), 1, 0, 'C')
                pdf.ln(col_size)
        else:
            pdf.cell(50, head_size, "Newly added rows", 0, 0, 'C')
            pdf.ln(head_size)
            column_list = data.columns.tolist()
            # formating the column names
            for col in column_list:
                pdf.cell(col_width, col_size, str(col), 1, 0, 'C')
            pdf.ln(col_size)
            # formating the data
            for ind in data.index.tolist():
                for col in column_list:
                    pdf.cell(col_width, col_size, str(
                        data[col].loc[ind]), 1, 0, 'C')
                pdf.ln(col_size)
    else:
        pdf.cell(50, 10, "No data changes", 0, 0, 'C')
    pdf.output(file_path, 'F')


# sending the mail

def send_mail(sender_mail, send_to,
              email_subject,
              email_message,
              attachment_location=''):
    msg = MIMEMultipart()
    msg['From'] = sender_mail
    msg['To'] = send_to
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))
    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(filename, 'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename=%s" % filename)
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.outlook.com', 587)
        server.ehlo()
        server.starttls()
        server.login(mycredits.login['email_id'], mycredits.login['password'])
        text = msg.as_string()
        server.sendmail(sender_mail, send_to, text)
        print("email sent")
        server.quit()
    except Exception as e:
        print("SMpt server connection error")
        exceptionbox(msg="SMPT server connection error",
                     title="Error while sending mail)

    return True


if __name__ == "__main__":
    _run = CompareModifiyData()
    data, data_type = _run.read_files()
    where_to_save = enterbox("enter your filename/path with .pdf extension:")
    create_pdf(data, data_type, where_to_save)
    send_mail(sender_mail=mycredits.login['email_id'],
              send_to="darawot709@btsese.com",
              email_subject="stmp test mail case1",
              email_message="sending mail from python code to \
                    check the smtp\n find the attached file",
              attachment_location=where_to_save)
