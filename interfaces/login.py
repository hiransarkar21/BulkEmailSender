from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *


# master login window
class LoginWindow(QWidget):
    def __init__(self):
        super(LoginWindow, self).__init__()

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
        self.login_screen_width = self.screen_size.width() // 2.7
        self.login_screen_height = self.screen_size.height() // 2

        # instance methods
        self.window_configurations()
        self.user_interface()

    def window_configurations(self):
        self.setFixedSize(int(self.login_screen_width), int(self.login_screen_height))
        self.setWindowTitle("Login into - Bulk Email Sender")

    def user_interface(self):
        # parent and their respective child layouts
        self.master_layout = QVBoxLayout()
        self.master_layout.setContentsMargins(10, 60, 10, 10)
        self.header_layout = QVBoxLayout()
        self.body_layout = QVBoxLayout()
        self.body_layout.setContentsMargins(40, 10, 40, 10)
        self.child_email_layout = QHBoxLayout()
        self.child_password_layout = QHBoxLayout()
        self.footer_layout = QHBoxLayout()

        # widgets
        self.welcome_label = QLabel()
        self.welcome_label.setFont(self.heading_font)
        self.welcome_label.setText("Welcome to Bulk Email Sender")

        self.modules_and_libraries_used_label = QLabel()
        self.modules_and_libraries_used_label.setFont(self.secondary_heading_font)
        self.modules_and_libraries_used_label.setText("PyQt5 | SMTP | Requests")

        self.email_address_label = QLabel()
        self.email_address_label.setFont(self.paragraph_font)
        self.email_address_label.setText("Email : ")

        self.get_email_address = QLineEdit()
        self.get_email_address.setFont(self.paragraph_font)
        self.get_email_address.setPlaceholderText(" email address ...")
        self.get_email_address.setFixedSize(int(self.width() // 1.7), int(self.height() // 10))
        self.get_email_address.setStyleSheet("""QLineEdit{border-radius: 20px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.password_label = QLabel()
        self.password_label.setFont(self.paragraph_font)
        self.password_label.setText("Password : ")

        self.get_user_password = QLineEdit()
        self.get_user_password.setFont(self.paragraph_font)
        self.get_user_password.setPlaceholderText(" password ...")
        self.get_user_password.setEchoMode(QLineEdit.Password)
        self.get_user_password.setFixedSize(int(self.width() // 1.7), int(self.height() // 10))
        self.get_user_password.setStyleSheet("""QLineEdit{border-radius: 20px; padding-right: 15px; 
                padding-left: 15px;}""")

        self.login_button = QPushButton()
        self.login_button.setFont(self.paragraph_font)
        self.login_button.setText(" Login ")
        self.login_button.setFixedSize(int(self.width() // 4), int(self.height() // 10))
        self.login_button.clicked.connect(self.open_bulk_email_sender_window)
        self.login_button.setStyleSheet("""QPushButton{border-radius: 25px; background-color: green; color: white;}""")

        # adding widgets to respective layouts
        self.header_layout.addWidget(self.welcome_label, alignment=Qt.AlignHCenter)
        self.header_layout.addWidget(self.modules_and_libraries_used_label, alignment=Qt.AlignHCenter)

        self.child_email_layout.addWidget(self.email_address_label)
        self.child_email_layout.addWidget(self.get_email_address)

        self.child_password_layout.addWidget(self.password_label)
        self.child_password_layout.addWidget(self.get_user_password)

        self.footer_layout.addWidget(self.login_button)

        # adding child email and password layout to parent body_layout
        self.body_layout.addLayout(self.child_email_layout)
        self.body_layout.addLayout(self.child_password_layout)

        # adding header, body and footer layout to master_layout
        self.master_layout.addLayout(self.header_layout)
        self.master_layout.addStretch()
        self.master_layout.addLayout(self.body_layout)
        self.master_layout.addStretch()
        self.master_layout.addLayout(self.footer_layout)
        self.master_layout.addStretch()

        self.setLayout(self.master_layout)

    def open_bulk_email_sender_window(self):
        from interfaces import bulk_email_sender

        self.bulk_email_sender_window = bulk_email_sender.EmailSender()
        self.bulk_email_sender_window.show()
        self.close()
