import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import re

def send_email(to_email, pdf_file_route, full_name, pdf_to_send):
  user_name = 'petar.klenovic@estudent.hr'
  passwd = 'Kamikaza36!'
  from_email = 'petar.klenovic@estudent.hr'
  #to_email = 'petar.klenovic@gmail.com'

  msg = MIMEMultipart()

  msg['From'] = from_email
  msg['To'] = to_email
  msg['Subject'] = 'eSTUDENT Certifikat'
  msg_text = "Dragi/a " + full_name + ',\n' + "iza nas je jedna uspješna i produktivna godina, u kojoj si uložio/la puno truda kako bi ostvario/la ciljeve svog tima, ali i cijele Udruge. Želimo ti se zahvaliti na velikom doprinosu te kao dokaz o tvojim ovogodišnjim aktivnostima i radu, šaljemo ti čuveni certifikat o kojem smo ove godine toliko pričali, koji možeš priložiti uz svoju prijavu za posao. Nadamo se da ćeš ovu godinu pamtiti po lijepim događajima i prijateljstvima koje si stekao/la.\n\nP. S. Molimo te da pregledaš svoj certifikat i povratno se javiš na ovaj mail ili na it.podrska@estudent.hr ako primijetiš neku grešku.\nUgodan dan!\n\n--\n"
  #msg_text = "Dragi/a " + full_name + ',\n' + "u prilogu ti šaljemo ispravljeni certifikat. Ako primjetiš još neku grešku slobodno se javi povratno na ovaj mail ili na it.podrska@estudent.hr \n" + "Tvoj IT tim."
  signature_html = """<div dir="ltr" class="gmail_signature" data-smartmail="gmail_signature"><div dir="ltr"><span><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(0,0,0);font-weight:700;vertical-align:baseline;white-space:pre-wrap">Petar Klenović</span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(0,0,0);vertical-align:baseline;white-space:pre-wrap">Voditelj, Informacijske tehnologije</span></p><br><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(153,153,153);vertical-align:baseline;white-space:pre-wrap">E-pošta:</span><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(34,34,34);vertical-align:baseline;white-space:pre-wrap"> </span><a href="mailto:ime.prezime@estudent.hr" target="_blank"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(17,85,204);vertical-align:baseline;white-space:pre-wrap">petar.klenovic@estudent.hr</span></a></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(153,153,153);vertical-align:baseline;white-space:pre-wrap">Mobitel:</span><span style="font-size:10pt;font-family:Poppins,sans-serif;vertical-align:baseline;white-space:pre-wrap"><font color="#222222"> <a href="tel:++385993379949" target="_blank">0993379949</a></font></span></p><br><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(204,0,0);font-weight:700;vertical-align:baseline;white-space:pre-wrap"><span style="border:none;display:inline-block;overflow:hidden;width:110px;height:24px"><img src="https://lh4.googleusercontent.com/pn7Ot6rr35k1ANTETeY0pclp9LUS5QnP9GMgDaa0iTwVQl6XRdGIH9zpXvgATgqBtA1NL0NQg87oUsPQ18mdk0b-_wRBmW87cVN1ZsWdHpQ0e6Z7_ae51URGHHKMd7jehuGsfPSl" width="110" height="24" style="margin-left:0px;margin-top:0px"></span></span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(34,34,34);vertical-align:baseline;white-space:pre-wrap">Trg J. F. Kennedyja 6, p.p. 137, Zagreb</span></p><p dir="ltr" style="line-height:1.38;margin-top:0pt;margin-bottom:0pt"><span style="font-size:10pt;font-family:Poppins,sans-serif;color:rgb(17,85,204);vertical-align:baseline;white-space:pre-wrap"><a href="https://www.estudent.hr/" target="_blank">www.estudent.hr</a></span></p><div><br></div></span></div></div>"""
  msg.attach( MIMEText(msg_text) )
  msg.attach(MIMEText(signature_html,'html'))

  part = MIMEBase('application', "octet-stream")
  fo=open(pdf_file_route,"rb")
  part.set_payload(fo.read() )
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', 'attachment', filename=pdf_to_send)
  msg.attach(part)

  server = smtplib.SMTP('smtp.gmail.com', 587)
  server.ehlo()
  server.starttls()
  server.login(user_name, passwd)

  server.sendmail(from_email, to_email, msg.as_string())
  server.close()
  print("Mail sent to: " + to_email)

cert_folder = os.listdir('./Certifikati-2022')
#print(cert_folder)
for team in cert_folder:
  team_members = os.listdir('./Certifikati-2022/' + team)
  #print(team_members)
  for member in team_members:
    to_email = member
    pdf_to_send = os.listdir('./Certifikati-2022/' + team + '/' + member)[0]
    #pdf_file_name = pdf_to_send[0]
    #print(pdf_to_send)
    pdf_file_route = './Certifikati-2022/' + team + '/' + member + '/' + pdf_to_send
    full_name = re.search('.+?(?=eSTUDENT)', pdf_to_send).group(0).strip()
    try:
      send_email(to_email, pdf_file_route, full_name, pdf_to_send)
      file2 = open("mail_sent.txt","a")
      file2.write(to_email + "; " + "\n")
      file2.close()
    except:
      print("Something went wrong! Email: " + to_email)
      print(os.error)
      file1 = open("mail_failed.txt","a")
      file1.write(to_email + "; " + "\n")
      file1.close()
