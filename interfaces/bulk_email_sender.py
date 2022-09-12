from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import smtplib


# bulk email sender window
class EmailSender(QWidget):
    def __init__(self):
        super(EmailSender, self).__init__()

        # global attributes
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

        # instance methods
        self.window_configurations()
        self.user_interface()

    def window_configurations(self):
        self.setFixedSize(int(self.bulk_email_sender_screen_width), int(self.bulk_email_sender_screen_height))
        self.setWindowTitle("Bulk Email Sender")

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
        self.email_address_label.setText("Email : ")

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
        self.get_email_password.setPlaceholderText("password ...")
        self.get_email_password.setEchoMode(QLineEdit.Password)
        self.get_email_password.setFixedSize(int(self.width() // 1.5), int(self.height() // 16))
        self.get_email_password.setStyleSheet("""QLineEdit{border-radius: 10px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.email_recipients_label = QLabel()
        self.email_recipients_label.setFont(self.paragraph_font)
        self.email_recipients_label.setText("Recipients : ")

        self.get_email_recipients_address = QLineEdit()
        self.get_email_recipients_address.setFont(self.paragraph_font)
        self.get_email_recipients_address.setPlaceholderText("recipients ...")
        self.get_email_recipients_address.setFixedSize(int(self.width() // 1.5), int(self.height() // 16))
        self.get_email_recipients_address.setStyleSheet("""QLineEdit{border-radius: 10px; padding-right: 15px; 
                padding-left: 15px;}""")

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
        self.get_email_body.setStyleSheet("""QTextEdit{border-radius: 20px; border: 0px;}""")

        self.email_options_groupbox = QGroupBox()
        self.email_options_groupbox.setFlat(True)
        self.email_options_groupbox.setStyleSheet("""QGroupBox{background-color: transparent; border: 2px solid silver; 
        border-radius: 10px;}""")

        self.email_options_label = QLabel()
        self.email_options_label.setText("Email Options : ")
        self.email_options_label.setFont(self.paragraph_font)

        self.sender_sleep_label = QLabel()
        self.sender_sleep_label.setFont(self.paragraph_font)
        self.sender_sleep_label.setText("Sleep For : ")

        self.get_sleep_period = QLineEdit()
        self.get_sleep_period.setFont(self.paragraph_font)
        self.get_sleep_period.setPlaceholderText("seconds")
        self.get_sleep_period.setFixedSize(int(self.width() // 6), int(self.height() // 19))
        self.get_sleep_period.setStyleSheet("QLineEdit{border-radius: 20px; padding-left: 10px; padding-right: 20px;}")

        self.send_email_button = QPushButton()
        self.send_email_button.setFont(self.paragraph_font)
        self.send_email_button.setText("Send!")
        self.send_email_button.setFixedSize(int(self.width() // 5), int(self.height() // 17))
        self.send_email_button.setStyleSheet("""QPushButton{border-radius: 25px; background-color: green; color: white;}""")

        # adding widgets to their layouts
        self.header_layout.addWidget(self.welcome_label, alignment=Qt.AlignHCenter)
        self.header_layout.addWidget(self.modules_and_libraries_used_label, alignment=Qt.AlignHCenter)

        self.child_email_address_layout.addWidget(self.email_address_label)
        self.child_email_address_layout.addWidget(self.get_email_address)

        self.child_email_password_layout.addWidget(self.email_password_label)
        self.child_email_password_layout.addWidget(self.get_email_password)

        self.child_email_recipients_layout.addWidget(self.email_recipients_label)
        self.child_email_recipients_layout.addWidget(self.get_email_recipients_address)

        self.child_email_subject_layout.addWidget(self.email_subject_label)
        self.child_email_subject_layout.addWidget(self.get_email_subject)

        self.child_email_body_layout.addWidget(self.email_body_label)
        self.child_email_body_layout.addWidget(self.get_email_body)

        self.child_email_sender_sleep_layout.addWidget(self.sender_sleep_label)
        self.child_email_sender_sleep_layout.addWidget(self.get_sleep_period)
        self.child_email_sender_sleep_layout.addStretch()

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


