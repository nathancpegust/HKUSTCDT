import smtplib
import pandas as pd

from email.message import EmailMessage

EMAIL_ADDRESS = 'hkustcdt@gmail.com'
EMAIL_PASSWORD = 'nakedwai'

e = pd.read_excel('info.xlsx')
Judges = e['Judges'].values
Emails = e['Emails'].values

for i in range(len(Judges)):
    CurrentEmail = Emails[i]
    msg = EmailMessage()
    msg['Subject'] = '誠邀擔任香港科技大學與香港城市大學辯論友誼賽之評審工作'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = CurrentEmail
    msg.add_alternative("""\
<!DOCTYPE html>
<html>
    <body>
        <p dir="ltr" style="line-height:1.2;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:12pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">{judge}尊鑒：</span></p>
<p dir="ltr" style="line-height:1.2;text-align: center;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:12pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:700;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">誠邀擔任香港科技大學與香港城市大學辯論友誼賽之評審工作</span></p>
<p dir="ltr" style="line-height:1.2;text-indent: 6pt;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">　一年一度學界盛事「大專辯論賽」將至，各院校亦相繼舉辦模擬辯論賽為大專盃作準備。香港科技大學特與香港城市大學合辦友誼辯論賽一場，籍此切磋辯技之餘，亦希望得到寶貴意見，改進自身不足。</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">　　現誠邀 &nbsp;閣下擔任是次友誼賽評判，為友賽進行評分工作。特附上有關比賽詳情：</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">日期：二零二一年一月二十三日（星期日）</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">地點： 荃灣西樓角路222-224號豪輝商業中心二座UG2 - 01室 康溢教育中心（評判可選擇線上或到場作評分工作）</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">時間：</span><span style="font-size:11pt;font-family:'Arial Black',sans-serif;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">&nbsp;</span><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">下午三時至五時</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">辯題： 零工經濟對香港勞動市場利大於弊</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">正方：香港科技大學</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">反方：香港城市大學</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;">&nbsp;</p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:12pt;font-family:PMingLiu;color:#313131;background-color:#ffffff;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">不知 閣下意下如何，</span><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">敬候示覆，如蒙應允，不勝感激。如有垂詢，煩請此電郵與本人聯絡。</span></p>
<p dir="ltr" style="line-height:1.2;text-indent: 22pt;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">敬祝</span></p>
<p dir="ltr" style="line-height:1.2;text-align: justify;background-color:#ffffff;margin-top:0pt;margin-bottom:10pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">教安</span></p>
<p dir="ltr" style="line-height:1.2;margin-right: 22pt;text-align: right;background-color:#ffffff;margin-top:0pt;margin-bottom:10pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">香港科技大學粵語辯論隊隊長　</span></p>
<p dir="ltr" style="line-height:1.2;text-align: right;background-color:#ffffff;margin-top:0pt;margin-bottom:10pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">　彭子健 謹啟</span></p>
<p dir="ltr" style="line-height:1.2;margin-right: 44pt;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;"><span style="font-size:11pt;font-family:PMingLiu;color:#000000;background-color:transparent;font-weight:400;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre;white-space:pre-wrap;">二零二二年一月十六日</span></p>
<p dir="ltr" style="line-height:1.2;margin-right: 44pt;background-color:#ffffff;margin-top:0pt;margin-bottom:0pt;padding:0pt 0pt 10pt 0pt;">&nbsp;</p>
<p dir="ltr" style="line-height:1.2;margin-right: 44pt;background-color:#ffffff;margin-top:0pt;margin-bottom:10pt;">&nbsp;</p>    </body>
</html>
""".format(judge=Judges[i]), subtype='html')

    


    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)