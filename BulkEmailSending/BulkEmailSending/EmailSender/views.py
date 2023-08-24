import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status


class Excel(APIView):
    def post(self, request):
        try:
            name=request.data.get("name")
            MultipleName=request.data.get("MultipleName")
            sheet=request.FILES.get("sheet")
            design=request.data.get("design")
            if sheet:
                html='''{design}'''
                a=pd.read_excel(sheet)
                for i in list(a):
                    r=(a[i].tolist())
                    for j in r:
                        def sendmailoforderdetailsCustomer(email_to="",design=''):
                            email_from = 'selnoxinfo@gmail.com'
                            password = 'ddylfolnferwhjue'
                            date_str = pd.Timestamp.today().strftime('%Y-%m-%d')
                            email_message = MIMEMultipart()
                            email_message['From'] = email_from
                            email_message['To']=email_to
                            email_message['Subject'] = f'promotion',date_str
                            email_message.attach(MIMEText(html.format(design=design),"html"))
                            email_string = email_message.as_string()
                            context = ssl.create_default_context()
                            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                                server.login(email_from, password)
                                server.sendmail(email_from, email_to, email_string)
                    sendmailoforderdetailsCustomer(email_to=j,design=design)
                return Response("Success", status=status.HTTP_200_OK)
            
            elif MultipleName:
                for i in MultipleName:
                    def sendmailoforderdetailsCustomer(email_to="",design=''):

                            email_from = 'selnoxinfo@gmail.com'
                            password = 'ddylfolnferwhjue'
                            date_str = pd.Timestamp.today().strftime('%Y-%m-%d')
                            email_message = MIMEMultipart()
                            email_message['From'] = email_from
                            email_message['To']=email_to
                            email_message['Subject'] = f'promotion',date_str
                            email_message.attach(MIMEText(html.format(design=design),"html"))
                            email_string = email_message.as_string()
                            context = ssl.create_default_context()
                            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                                server.login(email_from, password)
                                server.sendmail(email_from, email_to, email_string)
                sendmailoforderdetailsCustomer(email_to=i,design=design)
                return Response("Success", status=status.HTTP_200_OK)
            
            else:
                def sendmailoforderdetailsCustomer(email_to="",design=''):

                            email_from = 'selnoxinfo@gmail.com'
                            password = 'ddylfolnferwhjue'
                            date_str = pd.Timestamp.today().strftime('%Y-%m-%d')
                            email_message = MIMEMultipart()
                            email_message['From'] = email_from
                            email_message['To']=email_to
                            email_message['Subject'] = f'promotion',date_str
                            email_message.attach(MIMEText(html.format(design=design),"html"))
                            email_string = email_message.as_string()
                            context = ssl.create_default_context()
                            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                                server.login(email_from, password)
                                server.sendmail(email_from, email_to, email_string)
                sendmailoforderdetailsCustomer(email_to=name,design=design)
                return Response("Success", status=status.HTTP_200_OK)
                
        except Exception as e:
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)