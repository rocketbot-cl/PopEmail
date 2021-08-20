# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""

from email.parser import Parser
import os
import re
import sys


base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'PopEmail' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

class PopEmail:
    def __init__(self, username, password, host, port=110, ssl=False):
        import poplib
        self.HOST = host
        self.PORT = port
        self.connection = poplib.POP3_SSL(host, port) if ssl else poplib.POP3(host, port)
        self.connection.set_debuglevel(1)
        self.USERNAME = username
        self.PASSWORD = password
        self.connected = False
    def connect(self):
        try:
            print("user")
            self.connection.user(self.USERNAME)
            print("pass")
            self.connection.pass_(self.PASSWORD)

            self.connected = True
            return True
        except:
            PrintException()
            return False

    def close(self):
        self.connection.quit()
        self.connected = False
global pop_email
"""
    Obtengo el modulo que fue invocado
"""

module = GetParams("module")


def get_attachments(messages, folder):
        attachments = []
        try:
            for msg in messages:
                for part in msg.walk():
                    name = part.get_filename()
                    if name is None:
                        continue
                    data = part.get_payload(decode=True)
                    name = name.replace('\r\n', '')
                    with open(folder + os.sep + name,'wb') as f:
                        f.write(data)
                        f.close()
                    attachments.append(name)
        except Exception as e:
            PrintException()
        return attachments


def parse_uid(tmp):
    pattern_uid = re.compile('\d+ \(UID (?P<uid>\d+)\)')
    print('tmp', tmp)
    try:
        tmp = tmp.decode()
    except:
        pass
    match = pattern_uid.match(tmp)
    return match.group('uid')


if module == "conf_mail":

    conx = ""

    try:
        user = GetParams('user')
        pass_ = GetParams('password')
        host = GetParams("host")
        port = GetParams("port")
        ssl = GetParams("ssl")
        var_ = GetParams('var_')

        if ssl is None:
            ssl = "False"
        ssl = eval(ssl)
        pop_email = PopEmail(user, pass_, host, port=int(port), ssl=ssl)
        OK = pop_email.connect()

    except:
        PrintException()
        OK = False

    SetVar(var_, OK)


if module == "get_mail":
    filter_ = GetParams('filter')
    value = GetParams("value")
    var_ = GetParams('var_')

    try:
        if pop_email.connected:
            num_msg = pop_email.connection.stat()[0]
            mails_id = []
            for i in range(1,num_msg + 1):
                if not filter_:
                    mails_id.append(i)
                else:
                    res, lines, octets = pop_email.connection.retr(i)
                    msg_content = b'\r\n'.join(lines).decode('utf-8')
                    msg = Parser().parsestr(msg_content)
                    if value in msg.get(filter_):
                        mails_id.append(i)


            SetVar(var_, mails_id)

        else:
            raise Exception("Run server configuration command")

    except Exception as e:
        PrintException()
        raise e

if module == "read_mail":
    id_ = GetParams('id_')
    var_ = GetParams('var_')
    att_folder = GetParams('att_folder')

    try:
        if pop_email.connected:
            num_msg = pop_email.connection.stat()[0]
        
            if int(id_) > num_msg:
                raise Exception(f"There is not id {id_}")

            res, lines, octets = pop_email.connection.retr(int(id_))
            msg_content = b'\r\n'.join(lines).decode('utf-8')
            msg = Parser().parsestr(msg_content)
            print(dir(msg))
            att = get_attachments([msg], att_folder)
            mail = {
                "From": msg.get("From"),
                "To": msg.get("To"),
                "Subject": msg.get("Subject"),
                "Date": msg.get("Date"),
                "attachments": att
            }
            SetVar(var_, mail)
        else:
            raise Exception("Run server configuration command")


    except Exception as e:
        PrintException()
        raise e

if module == "close":
    if pop_email.connected:
        pop_email.close()

# if module == "get_unread":
#     filtro = GetParams('filtro')
#     var_ = GetParams('var_')
#
#     try:
#         mail = imaplib.IMAP4_SSL('imap.gmail.com')
#         mail.login(fromaddr, password)
#         mail.list()
#         # Out: list of "folders" aka labels in gmail.
#         mail.select("inbox")  # connect to inbox.
#
#         if filtro and len(filtro) > 0:
#             result, data = mail.search(None, filtro, "UNSEEN")
#         else:
#             result, data = mail.search(None, "UNSEEN")
#
#         ids = data[0]  # data is a list.
#         id_list = ids.split()  # ids is a space separated string
#
#         # print('ID',id_list)
#         lista = [b.decode() for b in id_list]
#
#         SetVar(var_, lista)
#     except Exception as e:
#         PrintException()
#         raise e


# if module == "reply_email":
#     id_ = GetParams('id_')
#     body_ = GetParams('body')
#     attached_file = GetParams('attached_file')
#     # print(body_, attached_file)
#
#     try:
#         mail = imaplib.IMAP4_SSL('imap.gmail.com')
#         mail.login(fromaddr, password)
#         mail.select("inbox")
#         server = smtplib.SMTP('smtp.gmail.com', 587)
#         server.starttls()
#         server.login(fromaddr, password)
#
#         # mail.select()
#         typ, data = mail.fetch(id_, '(RFC822)')
#         raw_email = data[0][1]
#         mm = email.message_from_bytes(raw_email)
#
#         # msg = MIMEMultipart()
#         # msg.attach(MIMEText(body_, 'plain'))
#
#         #    m_ = create_auto_reply(mm, body_)
#         mail__ = MIMEMultipart()
#         mail__['Message-ID'] = make_msgid()
#         mail__['References'] = mail__['In-Reply-To'] = mm['Message-ID']
#         mail__['Subject'] = 'Re: ' + mm['Subject']
#         mail__['From'] = mm['To'] = mm['Reply-To'] or mm['From']
#         mail__.attach(MIMEText(dedent(body_), 'html'))
#
#         if attached_file:
#             if os.path.exists(attached_file):
#                 filename = os.path.basename(attached_file)
#                 attachment = open(attached_file, "rb")
#                 part = MIMEBase('application', 'octet-stream')
#                 part.set_payload((attachment).read())
#                 attachment.close()
#                 encoders.encode_base64(part)
#                 part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
#                 mail__.attach(part)
#
#         # print("FROMADDR",fromaddr, "FROM",mm['From'], "TO:",mm['To'])
#         server.sendmail(fromaddr, mm['From'], mail__.as_bytes())
#         # server.sendmail(fromaddr, mm['To'], mail__.as_bytes())
#         # server.close()
#         mail.logout()
#     except Exception as e:
#         PrintException()
#         raise e
#
# if module == "create_folder":
#     try:
#         folder_name = GetParams('folder_name')
#         host = "imap.gmail.com"
#         mail = imaplib.IMAP4_SSL(host)
#         mail.login(fromaddr, password)
#         mail.create(folder_name)
#     except:
#         PrintException()
#         raise e
#
# if module == "move_mail":
#     # imap = GetGlobals('email')
#     id_ = GetParams("id_")
#     label_ = GetParams("label_")
#     var = GetParams("var")
#
#     if not id_:
#         raise Exception("No ha ingresado ID de email a mover")
#     if not label_:
#         raise Exception("No ha ingresado carpeta de destino")
#     try:
#         # login on IMAP server
#         # if imap.IMAP_SSL:
#         #     mail = imaplib.IMAP4_SSL('imap.gmail.com')
#         # else:
#         #     mail = imaplib.IMAP4('imap.gmail.com')
#
#         mail = imaplib.IMAP4_SSL('imap.gmail.com')
#         mail.login(fromaddr, password)
#         mail.select('inbox', readonly=False)
#         resp, data = mail.fetch(id_, "(UID)")
#         msg_uid = parse_uid(data[0])
#
#         result = mail.uid('COPY', int(msg_uid), label_)
#
#         if result[0] == 'OK':
#             mov, data = mail.uid('STORE', msg_uid, '+FLAGS', '(\Deleted)')
#             res = mail.expunge()
#             if var:
#                 ret = True if res[0] == 'OK' else False
#                 SetVar(var, ret)
#         else:
#             raise Exception(result)
#     except Exception as e:
#         PrintException()
#         raise e
#
# if module == "markAsUnread":
#     id_ = GetParams("id_")
#     var = GetParams("var")
#
#     try:
#         mail = imaplib.IMAP4_SSL('imap.gmail.com')
#         mail.login(fromaddr, password)
#         mail.select('inbox', readonly=False)
#         resp, data = mail.fetch(id_, "(UID)")
#         msg_uid = parse_uid(data[0])
#
#         data = mail.uid('STORE', msg_uid, '-FLAGS', '(\Seen)')
#     except Exception as e:
#         PrintException()
#         raise e
#
#
#
