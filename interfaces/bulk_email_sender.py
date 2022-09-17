from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from openpyxl import load_workbook
import smtplib
import json
import time
import os

BASE_DIR = os.getcwd()
APPLICATION_DATA = os.path.join(BASE_DIR, "application_data")


class SMTPThread(QThread):
    def __init__(self, email_address, password, metadata_file_location):
        super(SMTPThread, self).__init__()

        # global attributes
        self.email_address = email_address
        self.password = password
        self.metadata_file_location = metadata_file_location
        self.is_thread_active = True

    def run(self):
        with open(self.metadata_file_location) as metadata_file:
            metadata = json.load(metadata_file)

        print("Preparing to send emails ...", end="\n")
        sleep_time = metadata["sleep_time"]
        smtp_host = metadata["smtp_host"]
        smtp_port = metadata["smtp_port"]
        recipients_excel_sheet = metadata["recipients_excel_sheet"]
        email_subject = metadata["subject"]
        email_body = metadata["body"]

        # fetching data from Excel sheet
        self.master_workbook = load_workbook(filename=recipients_excel_sheet)
        self.master_sheet = self.master_workbook.active
        recipients_addresses = [excel_object.value for excel_object in self.master_sheet["A"] if
                                excel_object.value != "Email Addresses"]

        while self.is_thread_active:
            # looping through recipients addresses
            for recipient in recipients_addresses:
                header = 'To: ' + ", " + recipient + '\n' + 'From: ' + self.email_address + \
                         '\n' + 'Subject: ' + email_subject + '\n' + \
                         'List-Unsubscribe: ' + 'mailto: unsubrequests@mtbservices.org?subject=unsubscribe' + ''

                final_message = header + '\n' + '\n' + email_body + '\n\n'

                # Sending Process
                server = smtplib.SMTP(smtp_host, smtp_port)
                server.starttls()
                server.login(self.email_address, self.password)
                server.sendmail(recipient, recipient, final_message)

                # testing console log
                print("Message sent to : ", recipient)
                time.sleep(sleep_time)

            server.quit()
            print("Complete sending emails ...")
            return


# bulk email sender window
class EmailSender(QWidget):
    def __init__(self):
        super(EmailSender, self).__init__()

        # global attributes
        self.smtp_thread = None

        self.heading_font = QFont("Poppins", 18)
        self.heading_font.setBold(True)
        self.heading_font.setWordSpacing(2)
        self.heading_font.setLetterSpacing(QFont.AbsoluteSpacing, 1)

        self.secondary_heading_font = QFont("Poppins", 16)
        self.secondary_heading_font.setWordSpacing(2)
        self.secondary_heading_font.setLetterSpacing(QFont.AbsoluteSpacing, 1)

        self.paragraph_font = QFont("Poppins", 13)
        self.paragraph_font.setWordSpacing(2)
        self.paragraph_font.setLetterSpacing(QFont.AbsoluteSpacing, 1)

        self.screen_size = QApplication.primaryScreen().availableSize()
        self.bulk_email_sender_screen_width = self.screen_size.width() // 2.5
        self.bulk_email_sender_screen_height = self.screen_size.height() // 1.2

        self.email_metadata_file = None
        self.recipients_excel_sheet_file_name = None

        # instance methods
        self.window_configurations()
        self.user_interface()
        self.file_configurations()

    def window_configurations(self):
        self.setFixedSize(int(self.bulk_email_sender_screen_width), int(self.bulk_email_sender_screen_height))
        self.setWindowTitle("Bulk Email Sender")

    def file_configurations(self):
        if os.path.isdir(APPLICATION_DATA):
            pass
        else:
            os.mkdir(APPLICATION_DATA)

        self.email_metadata_file = os.path.join(APPLICATION_DATA, "email_metadata.json")

        # checking if metadata file exists
        if os.path.isfile(self.email_metadata_file):
            with open(self.email_metadata_file) as metadata_file:
                metadata = json.load(metadata_file)

            sleep_time = metadata["sleep_time"]
            smtp_host = metadata["smtp_host"]
            smtp_port = metadata["smtp_port"]

            # filling up non-priority fields
            self.get_smtp_port.setText(str(smtp_port))
            self.get_smtp_host.setText(smtp_host)
            self.get_sleep_period.setText(str(sleep_time))

        else:
            pass

    def user_interface(self):
        self.qlineedit_size = QSize(int(self.width() // 1.5), int(self.height() // 19))

        self.master_layout = QVBoxLayout()
        self.master_layout.setContentsMargins(50, 50, 50, 20)
        self.header_layout = QVBoxLayout()
        self.body_layout = QVBoxLayout()
        self.email_layout = QVBoxLayout()
        self.email_options = QVBoxLayout()
        self.child_email_address_layout = QHBoxLayout()
        self.child_email_password_layout = QHBoxLayout()
        self.child_email_subject_layout = QHBoxLayout()
        self.child_email_recipients_layout = QHBoxLayout()
        self.child_email_body_layout = QHBoxLayout()
        self.child_email_sender_sleep_layout = QHBoxLayout()
        self.child_host_and_port_layout = QHBoxLayout()
        self.footer_layout = QHBoxLayout()

        # widgets
        self.welcome_label = QLabel()
        self.welcome_label.setFont(self.heading_font)
        self.welcome_label.setText("Welcome to Bulk Email Sender")

        self.modules_and_libraries_used_label = QLabel()
        self.modules_and_libraries_used_label.setFont(self.secondary_heading_font)
        self.modules_and_libraries_used_label.setText("PyQt5 | SMTP | Requests")

        self.email_groupbox = QGroupBox()
        self.email_groupbox.setFlat(True)
        self.email_groupbox.setStyleSheet("""QGroupBox{background-color: transparent; border: 2px solid silver; 
        border-radius: 10px;}""")

        self.email_fields_labels = QLabel()
        self.email_fields_labels.setFont(self.paragraph_font)
        self.email_fields_labels.setText("Email Fields : ")

        self.email_address_label = QLabel()
        self.email_address_label.setFont(self.paragraph_font)
        self.email_address_label.setText("SMTP Email : ")

        self.get_email_address = QLineEdit()
        self.get_email_address.setFont(self.paragraph_font)
        self.get_email_address.setFixedSize(self.qlineedit_size)
        self.get_email_address.setPlaceholderText("email address ... ")
        self.get_email_address.setStyleSheet("""QLineEdit{border-radius: 10px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.email_password_label = QLabel()
        self.email_password_label.setFont(self.paragraph_font)
        self.email_password_label.setText("Password : ")

        self.get_email_password = QLineEdit()
        self.get_email_password.setFont(self.paragraph_font)
        self.get_email_password.setPlaceholderText("smtp password ...")
        self.get_email_password.setEchoMode(QLineEdit.Password)
        self.get_email_password.setFixedSize(int(self.width() // 1.5), int(self.height() // 16))
        self.get_email_password.setStyleSheet("""QLineEdit{border-radius: 10px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.email_recipients_label = QLabel()
        self.email_recipients_label.setFont(self.paragraph_font)
        self.email_recipients_label.setText("Recipients : ")

        self.get_email_recipients_address = QLineEdit()
        self.get_email_recipients_address.setFont(self.paragraph_font)
        self.get_email_recipients_address.setPlaceholderText("recipients excel sheet ...")
        self.get_email_recipients_address.setReadOnly(True)
        self.get_email_recipients_address.setFixedSize(int(self.width() // 2.02), int(self.height() // 16))
        self.get_email_recipients_address.setStyleSheet("""QLineEdit{border-radius: 10px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.recipients_excel_sheet_browse_button = QPushButton()
        self.recipients_excel_sheet_browse_button.setFont(self.paragraph_font)
        self.recipients_excel_sheet_browse_button.setText("Browse")
        self.recipients_excel_sheet_browse_button.clicked.connect(self.recipients_excel_sheet_browse_button_clicked)
        self.recipients_excel_sheet_browse_button.setFixedSize(int(self.width() // 6), int(self.height() // 16))
        self.recipients_excel_sheet_browse_button.setStyleSheet("""QPushButton{border-radius: 25px; 
        background-color: #041a3d; color: white;}""")

        self.email_subject_label = QLabel()
        self.email_subject_label.setFont(self.paragraph_font)
        self.email_subject_label.setText("Subject : ")

        self.get_email_subject = QLineEdit()
        self.get_email_subject.setFont(self.paragraph_font)
        self.get_email_subject.setPlaceholderText("subject ...")
        self.get_email_subject.setFixedSize(int(self.width() // 1.5), int(self.height() // 16))
        self.get_email_subject.setStyleSheet("""QLineEdit{border-radius: 10px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.email_body_label = QLabel()
        self.email_body_label.setFont(self.paragraph_font)
        self.email_body_label.setText("Message : ")

        self.get_email_body = QTextEdit()
        self.get_email_body.setFont(self.paragraph_font)
        self.get_email_body.setFixedSize(int(self.width() // 1.5), int(self.height() // 8))
        self.get_email_body.setPlaceholderText(" enter email body ...")
        self.get_email_body.setStyleSheet("""QTextEdit{border: 0px; border-radius: 20px;}""")

        self.email_options_groupbox = QGroupBox()
        self.email_options_groupbox.setFlat(True)
        self.email_options_groupbox.setStyleSheet("""QGroupBox{background-color: transparent; border: 2px solid silver; 
        border-radius: 10px;}""")

        self.email_options_label = QLabel()
        self.email_options_label.setText("Email Options : ")
        self.email_options_label.setFont(self.paragraph_font)

        self.sender_sleep_label = QLabel()
        self.sender_sleep_label.setFont(self.paragraph_font)
        self.sender_sleep_label.setText("Sleep For  : ")

        self.get_sleep_period = QLineEdit()
        self.get_sleep_period.setFont(self.paragraph_font)
        self.get_sleep_period.setPlaceholderText("seconds")
        self.get_sleep_period.setFixedSize(int(self.width() // 6), int(self.height() // 19))
        self.get_sleep_period.setStyleSheet("QLineEdit{border-radius: 20px; padding-left: 10px; padding-right: 20px;}")

        self.host_label = QLabel()
        self.host_label.setFont(self.paragraph_font)
        self.host_label.setText("SMTP Host : ")

        self.get_smtp_host = QLineEdit()
        self.get_smtp_host.setFont(self.paragraph_font)
        self.get_smtp_host.setPlaceholderText("host")
        self.get_smtp_host.setFixedSize(int(self.width() // 3.5), int(self.height() // 19))
        self.get_smtp_host.setStyleSheet("QLineEdit{border-radius: 20px; padding-left: 10px; padding-right: 20px;}")

        self.port_label = QLabel()
        self.port_label.setFont(self.paragraph_font)
        self.port_label.setText("Port : ")

        self.get_smtp_port = QLineEdit()
        self.get_smtp_port.setFont(self.paragraph_font)
        self.get_smtp_port.setPlaceholderText("port")
        self.get_smtp_port.setFixedSize(int(self.width() // 6), int(self.height() // 19))
        self.get_smtp_port.setStyleSheet("QLineEdit{border-radius: 20px; padding-left: 10px; padding-right: 20px;}")

        self.send_email_button = QPushButton()
        self.send_email_button.setFont(self.paragraph_font)
        self.send_email_button.setText("Send!")
        self.send_email_button.clicked.connect(self.send_email_button_clicked)
        self.send_email_button.setFixedSize(int(self.width() // 5), int(self.height() // 17))
        self.send_email_button.setStyleSheet(
            """QPushButton{border-radius: 25px; background-color: green; color: white;}""")

        # adding widgets to their layouts
        self.header_layout.addWidget(self.welcome_label, alignment=Qt.AlignHCenter)
        self.header_layout.addWidget(self.modules_and_libraries_used_label, alignment=Qt.AlignHCenter)

        self.child_email_address_layout.addWidget(self.email_address_label)
        self.child_email_address_layout.addWidget(self.get_email_address)

        self.child_email_password_layout.addWidget(self.email_password_label)
        self.child_email_password_layout.addWidget(self.get_email_password)

        self.child_email_recipients_layout.addWidget(self.email_recipients_label)
        self.child_email_recipients_layout.addWidget(self.get_email_recipients_address)
        self.child_email_recipients_layout.addWidget(self.recipients_excel_sheet_browse_button)

        self.child_email_subject_layout.addWidget(self.email_subject_label)
        self.child_email_subject_layout.addWidget(self.get_email_subject)

        self.child_email_body_layout.addWidget(self.email_body_label)
        self.child_email_body_layout.addWidget(self.get_email_body)

        self.child_email_sender_sleep_layout.addWidget(self.sender_sleep_label)
        self.child_email_sender_sleep_layout.addWidget(self.get_sleep_period)
        self.child_email_sender_sleep_layout.addStretch()

        self.child_host_and_port_layout.addWidget(self.host_label)
        self.child_host_and_port_layout.addWidget(self.get_smtp_host)
        self.child_host_and_port_layout.addSpacing(60)
        self.child_host_and_port_layout.addWidget(self.port_label)
        self.child_host_and_port_layout.addWidget(self.get_smtp_port)
        self.child_host_and_port_layout.addStretch()

        # adding child body layouts to master body layout
        self.email_layout.addSpacing(3)
        self.email_layout.addLayout(self.child_email_address_layout)
        self.email_layout.addLayout(self.child_email_password_layout)
        self.email_layout.addLayout(self.child_email_subject_layout)
        self.email_layout.addLayout(self.child_email_recipients_layout)
        self.email_layout.addLayout(self.child_email_body_layout)
        self.email_layout.addSpacing(3)
        self.email_groupbox.setLayout(self.email_layout)

        self.email_options.addSpacing(3)
        self.email_options.addLayout(self.child_email_sender_sleep_layout)
        self.email_options.addLayout(self.child_host_and_port_layout)
        self.email_options.addSpacing(3)
        self.email_options_groupbox.setLayout(self.email_options)

        self.body_layout.addWidget(self.email_fields_labels)
        self.body_layout.addSpacing(10)
        self.body_layout.addWidget(self.email_groupbox)
        self.body_layout.addSpacing(30)
        self.body_layout.addWidget(self.email_options_label)
        self.body_layout.addSpacing(10)
        self.body_layout.addWidget(self.email_options_groupbox)
        self.footer_layout.addWidget(self.send_email_button, alignment=Qt.AlignHCenter)

        # adding header, body and footer layout to master_layout
        self.master_layout.addLayout(self.header_layout)
        self.master_layout.addStretch()
        self.master_layout.addLayout(self.body_layout)
        self.master_layout.addStretch()
        self.master_layout.addLayout(self.footer_layout)

        self.setLayout(self.master_layout)

    def recipients_excel_sheet_browse_button_clicked(self):
        # opening open file dialog to get the location of the target Excel sheet
        excel_sheet_filename = QFileDialog.getOpenFileName(self, "Open Recipients Excel Sheet", "",
                                                           "Excel Sheet (*.xlsx)")

        self.recipients_excel_sheet_file_name = excel_sheet_filename[0]

        # setting location of target file to self.get_email_recipients_address
        self.get_email_recipients_address.setText(self.recipients_excel_sheet_file_name)

    def send_email_button_clicked(self):
        email_address = self.get_email_address.text()
        password = self.get_email_password.text()

        # writing information
        with open(self.email_metadata_file, "w") as email_metadata_file:
            metadata = {"sleep_time": int(self.get_sleep_period.text()),
                        "smtp_host": self.get_smtp_host.text(),
                        "smtp_port": int(self.get_smtp_port.text()),
                        "recipients_excel_sheet": self.recipients_excel_sheet_file_name,
                        "subject": self.get_email_subject.text(),
                        "body": self.get_email_body.toPlainText()}

            json.dump(metadata, email_metadata_file)

        # starting thread
        self.smtp_thread = SMTPThread(email_address=email_address, password=password,
                                      metadata_file_location=self.email_metadata_file)

        self.smtp_thread.start()
