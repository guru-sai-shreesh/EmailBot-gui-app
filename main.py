import smtplib
import speech_recognition as sr
import pyttsx3
from email.message import EmailMessage
import openpyxl as xl
from kivymd.app import MDApp
from kivy.core.window import Window
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivymd.uix.list import TwoLineAvatarListItem, IconLeftWidget
import copy

email_receivers = []
addresses = []
sub = []
body = []

screen_helper = """
ScreenManager:
    MenuScreen:
    SelectScreen:
    SubjectScreen:
    BodyScreen:
    EndScreen:


<MenuScreen>:
    Image:
        source: 'bot_anim.gif'
        anim_delay: 0.03
        mipmap: True
        allow_stretch: True
        pos_hint: {'center_x':0.5,'center_y':0.5}
    MDLabel:
        id: head
        text: 'EMAIL BOT'
        bold: True
        halign: 'center'
        bold: True
        font_size: '20sp'
        pos_hint: {'center_x':0.34,'center_y':0.83}
        color: 0.12, 0.45, 0.41, 1
    MDRectangleFlatButton:
        text: 'Start'
        pos_hint: {'center_x':0.75,'center_y':0.18}
        on_press: root.manager.current = 'select'

<SelectScreen>:
    name: 'select'
    BoxLayout:
        ScrollView:
            MDList:
                id: scroll
    MDRaisedButton:
        text: 'Speak names to add'
        pos_hint: {'center_x':0.58,'center_y':0.05}
        md_bg_color: app.theme_cls.primary_dark
        elevation: 12
        on_press: root.receiver_addresses()
    MDRectangleFlatButton:
        text: 'Next'
        pos_hint: {'center_x':0.88,'center_y':0.05}
        on_press: root.manager.current = 'subject'

<SubjectScreen>:
    name: 'subject'
    MDLabel:
        text: 'Subject:'
        bold: True
        halign: 'center'
        pos_hint: {'center_x':0.13,'center_y':0.95}
    MDRectangleFlatButton:
        id: sub
        text: "Press on 'start listening' to speak subject"
        pos_hint: {'center_x':0.5,'center_y':0.5}
        size_hint: 0.9, 0.8

    MDRaisedButton:
        text: 'Start listening'
        pos_hint: {'center_x':0.8,'center_y':0.95}
        md_bg_color: app.theme_cls.primary_dark
        on_press: root.listen_subject()

    MDRaisedButton:
        text: 'Next'
        pos_hint: {'center_x':0.87,'center_y':0.05}
        on_press: root.manager.current = 'body'

<BodyScreen>:
    name: 'body'
    MDLabel:
        text: 'Email Body: '
        bold: True
        halign: 'center'
        pos_hint: {'center_x':0.18,'center_y':0.95}
    MDRectangleFlatButton:
        id: ebod
        text: "Press on 'start listening' to speak BODY"
        pos_hint: {'center_x':0.5,'center_y':0.5}
        size_hint: 0.9, 0.8

    MDRaisedButton:
        text: 'Start listening'
        pos_hint: {'center_x':0.8,'center_y':0.95}
        md_bg_color: app.theme_cls.primary_dark
        on_press: root.listen_body()

    MDRaisedButton:
        text: 'Next'
        pos_hint: {'center_x':0.87,'center_y':0.05}
        on_press: root.manager.current = 'end'

<EndScreen>:
    name: 'end'
    MDRectangleFlatButton:
        text: "Press on 'Send' button to send composed Email"
        pos_hint: {'center_x':0.5,'center_y':0.7}
    MDRaisedButton:
        text: 'Send'
        pos_hint: {'center_x':0.5,'center_y':0.6}
        on_press: root.final_send()
    MDLabel:
        id: final
        halign: 'center'
        pos_hint: {'center_x':0.5,'center_y':0.5}
    MDRectangleFlatButton:
        id: notify
        pos_hint: {'center_x':0.47,'center_y':0.05}
    MDRaisedButton:
        id: back
        pos_hint: {'center_x':0.87,'center_y':0.05}
        on_press: root.manager.current = 'menu'
"""


class MenuScreen(Screen):
    pass


class SelectScreen(Screen):

    @staticmethod
    def receiver_addresses():
        receivers_str = mike_out()
        receivers = receivers_str.split(' and ')
        for receiver in receivers:
            if receiver not in contact_list:
                receivers.remove(receiver)
        if len(receivers) == 0:
            talk("sorry, there are no similar email addresses in your contacts")
            talk("Please try again")

        for receiver in receivers:
            email_receivers.append(receiver.capitalize())
            addresses.append(contact_list[receiver])
        print("Receiver's Email addresses: ", *addresses)


class SubjectScreen(Screen):
    line_count = 0

    def listen_subject(self):
        sub_line = mike_out()
        sub.append(sub_line)
        dup = copy.deepcopy(sub_line)
        c = 0
        for index in range(len(dup) - 1):
            if dup[index] == " " and c < 7:
                c += 1
            elif dup[index] == " " and c >= 7:
                dup = dup[:index] + '\n' + dup[index + 1:]
                c = 0
        if self.line_count == 0:
            self.ids.sub.text = dup.capitalize()
            self.line_count += 1
        else:
            self.ids.sub.text += f"\n{dup.capitalize()}"


class BodyScreen(Screen):
    line_count = 0

    def listen_body(self):
        body_line = mike_out()
        body.append(body_line)
        dup = copy.deepcopy(body_line)
        c = 0
        for index in range(len(dup) - 1):
            if dup[index] == " " and c < 7:
                c += 1
            elif dup[index] == " " and c >= 7:
                dup = dup[:index] + '\n' + dup[index + 1:]
                c = 0
        if self.line_count == 0:
            self.ids.ebod.text = dup.capitalize()
            self.line_count += 1
        else:
            self.ids.ebod.text += f"\n{dup.capitalize()}"


class EndScreen(Screen):

    def final_send(self):
        gather_and_send()
        final_msg_names = ""
        unique_receivers = list(set(email_receivers))
        for receiver in unique_receivers:
            final_msg_names += (' ' + receiver)
        self.ids.final.text = "Email(s) successfully sent to\n" + final_msg_names
        self.ids.notify.text = "To send New Email Press here->"
        self.ids.back.text = "Back"

    pass


# Create the screen manager
sm = ScreenManager()
sm.add_widget(MenuScreen(name='menu'))
sm.add_widget(SelectScreen(name='select'))
sm.add_widget(SubjectScreen(name='subject'))
sm.add_widget(BodyScreen(name='body'))
sm.add_widget(EndScreen(name='end'))

Window.size = (360, 600)

wb = xl.load_workbook('contacts.xlsx')
sheet = wb['Sheet1']
contact_list = {}
x = 2
for x in range(2, sheet.max_row):
    cell1 = sheet.cell(x, 2)
    cell2 = sheet.cell(x, 3)
    contact_list[cell1.value] = cell2.value

listener = sr.Recognizer()

engine = pyttsx3.init()
engine.setProperty('rate', 180)  # reduces WPM to 180
engine.setProperty('voice', 'com.apple.speech.synthesis.voice.samantha.premium')  # This voice is best suited in macos


def talk(text):
    engine.say(text)
    engine.runAndWait()


def mike_out():
    try:
        with sr.Microphone() as source:
            print('listening...')
            voice = listener.listen(source)
            info = listener.recognize_google(voice)
            print(info)
            return info.lower()
    except:
        pass


def send_email(receiver, subject, message):
    server = smtplib.SMTP('smtp.gmail.com', 587)  # to connect to server
    server.starttls()  # tells that it is an secure connection
    # Make sure to give app access in your Google account
    server.login('gojo.testing123@gmail.com', 'hellogojo')
    email = EmailMessage()
    email['From'] = 'gojo.testing123@gmail.com'
    email['To'] = receiver
    email['Subject'] = subject
    email.set_content(message)
    server.send_message(email)


def gather_and_send():
    Subject = ""
    Body = ""
    i = 0
    j = 0
    for subs in sub:
        if i == 0:
            Subject = subs.capitalize()
            i += 1
        else:
            Subject += ('. ' + subs.capitalize())
    for bodies in body:
        if j == 0:
            Body = bodies.capitalize()
            j += 1
        else:
            Body += ('. ' + bodies.capitalize())
    unique_addresses = list(set(addresses))
    for receiver in unique_addresses:
        send_email(receiver, Subject, Body)
    print('Email sent successfully to', *unique_addresses)


def new_contact():
    talk('type the name and address of new contact!')
    name = input("Type name of new contact to add: ")
    email = input("Type Email address of new contact to add: ")
    contact_list[name] = email
    new_cell0 = sheet.cell(x + 1, 1)
    new_cell1 = sheet.cell(x + 1, 2)
    new_cell2 = sheet.cell(x + 1, 3)
    new_cell0.value = x - 1
    new_cell1.value = name
    new_cell2.value = email
    wb.save('contacts.xlsx')
    wb.close()


class DemoApp(MDApp):
    def build(self):
        self.theme_cls.primary_palette = "Blue"
        screen = Screen()
        self.help_str = Builder.load_string(screen_helper)
        screen.add_widget(self.help_str)
        for key in contact_list:
            icons = IconLeftWidget(icon="contact_icon.jpeg")
            items = TwoLineAvatarListItem(text=key.capitalize(), secondary_text=contact_list[key])
            items.add_widget(icons)
            self.help_str.get_screen('select').ids.scroll.add_widget(items)
        return screen


DemoApp().run()
