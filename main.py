"""
A program to web scrap the assembla's website in order to automatically generate the invoices
"""

import requests
import urllib.request
import datetime
import bs4
import locale
import os
import email, smtplib, ssl
from openpyxl import Workbook, load_workbook
from bs4 import BeautifulSoup
import json
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



def main():

    config = {}
    try:
        with open("config.json", "r", encoding="utf-8") as configFile:
            config = json.load(configFile)
            iDate = [int(x) for x in config["initialDate"].split("/")]
            fDate = [int(x) for x in config["finalDate"].split("/")]

            config["initialDate"], config["finalDate"] = datetime.datetime(
            iDate[2], iDate[1], iDate[0]), datetime.datetime(fDate[2], fDate[1], fDate[0])

            config["sEmail"] = {"Y": True, "N": False}[
                config["sEmail"].upper()]
    except:
        print("File not found, answer the following prompts")
        config["username"] = input("Please enter your assembla's username ")
        config["password"] = input("Please enter your assembla's password ")
        config["name"] = input("Please enter your name ")
        config["path"] = input(
            "Please enter the path to store the invoice. Enter a dot for the current dir ")
        iDate = [int(x) for x in input(
            "Please enter the inital date for the invoice (dd/mm/yyyy) ").split("/")]
        fDate = [int(x) for x in input(
            "Please enter the final date for the invoice (dd/mm/yyyy) ").split("/")]
        
        config["initialDate"], config["finalDate"] = datetime.datetime(
            iDate[2], iDate[1], iDate[0]), datetime.datetime(fDate[2], fDate[1], fDate[0])

        config["sEmail"] = {"Y": True, "N": False}[
            input("Do you want to send an email with the generated invoice? Y/N ").upper()]

    finally:

        loginSession = loginIntoAssembla(
            config["username"], config["password"])

        if(checkIfLoggedIn(loginSession)):
            print("Logged succesfully")
            reports = getReportsFromDateRange(
                loginSession, config["username"], config["initialDate"], config["finalDate"])

            if(reports):
                filePath = modifyTemplate(config["name"], reports, config["path"])

                if(config["sEmail"]):
                    if("emailAddress" not in config.keys()):
                        config["emailAddress"] = input("Please enter your GMAIL address ")
                        config["emailPass"] = input("Please enter your GMAIL password ")
                        config["receiver"] = input("Please enter the recipient address ")

                    email = buildEmail(config["emailAddress"], config["receiver"],
                                    config["initialDate"], config["finalDate"], filePath)

                    sendEmail(config["emailAddress"], config["emailPass"], config["receiver"], email)
            else:
                print("There are no reports :(")
            print("Done :D")
        else:
            print("Log in failed. Please check your credentials and try again")

def getAuthToken(form):
    authInput = form.find(attrs = {"name": "authenticity_token"})
    # print("Auth input", authInput.attrs)
    return authInput['value']

def loginIntoAssembla(username, password):
    loginURL = "https://app.assembla.com/login"
    authURL = "https://auth.assembla.com/users/authenticate"

    session = requests.session()

    loginResponse = session.get(loginURL)

    soup = BeautifulSoup(loginResponse.text, "html.parser")
    form = soup.form

    payload = {
        "user[login]": username,
        "user[password]": password,
        "portfolio_id": '',
        "user[timezone]": getTimezone(form),
        "cookie[remember_me]": "on",
        "authenticity_token": getAuthToken(form),
        "commit": "Log in",
        "utf8": "✓"
    }

    loginResult = session.post(
        authURL, data=payload, headers=dict(referer=loginURL))
    # print("Login result", loginResult)

    # if("security_token" in session.cookies):
    #     print("Login cookies", session.cookies['security_token'])
    # else:
    #     print("Log in failed")

    return session


def modifyTemplate(name, reports, path):
    locale.setlocale(locale.LC_TIME, 'en-us')
    workbook = load_workbook("template.xlsx")
    worksheet = workbook.active
    
    startingRow = 6
    for i, report in enumerate(reports):
        # print("report", report, i)
        row = worksheet[startingRow + i]
        # print(row)
        row[0].value = report['date']
        if(i < len(reports) - 1):
            row[1].value = reports[i+1]['yesterday']
        else:
            row[1].value = report['today']
        
        row[2].value = report['startTime']
        row[3].value = report['hours']

    worksheet["C2"] = reports[0]['date'].strftime("%B %d")
    worksheet["D2"] = reports[-1]['date'].strftime("%B %d")
    worksheet["B3"] = name
    
    filename = reports[0]['date'].strftime(
        "%b %d")+" - "+reports[-1]['date'].strftime("%b %d")
    filePath = os.path.join(path, "{}.xlsx".format(filename))

    workbook.save(filePath)
    return filePath

def getReportFromDate(loginSession, user, date):
    sDate = date.strftime("%Y-%m-%d")
    standUpURL = "https://app.assembla.com/spaces/codigo-4-0/scrum"

    standUpResponse = loginSession.get(standUpURL, params = {"current_date": sDate})

    dailyReports = {
        "yesterday": None,
        "today": None,
        "date": date,
        "hours": 8,
        "startTime": "8:00"
    }
    # print("standUpResponse", standUpResponse)
    if(standUpResponse):
        standUpPage = BeautifulSoup(standUpResponse.text, 'html.parser')
        
        dataPanel = standUpPage.find(attrs={"data-panel": user.lower()})
        if(dataPanel):
            reportContainer = list(filter(lambda e: isinstance(
                e, bs4.element.Tag), dataPanel.parent.next_siblings))[0]
            headers = reportContainer.find_all('h4')
            containerContents = [x.next_sibling.next_sibling for x in headers]

            reports = []
            for cc in containerContents:
                # print(cc.contents)
                reports.append([])
                for content in cc:
                    if(isinstance(content, str)):
                        reports[-1].append(content)
                    elif(content.name == "br"):
                        reports[-1].append("\n")
                    elif(content.name == "a"):
                        reports[-1].append(content.string)
            # reports = [list(x.next_sibling.next_sibling.strings)
            #            for x in headers]

            # print("\n\n\n\n")
            # print(reports)

            dailyReports["yesterday"] = ''.join(reports[0])
            dailyReports["today"] = ''.join(reports[1])


        # print("The standup page content", dailyReports)

    
    return dailyReports

def getReportsFromDateRange(loginSession, user, initialDate, finalDate):
    dates = []
    while initialDate <= finalDate:
        dates.append(initialDate)
        initialDate+=datetime.timedelta(days=1)
    
    reports = [r for r in [getReportFromDate(loginSession, user, date) for date in dates] if r['yesterday'] or r['today']]
    return reports

def sendEmail(user, password, receiver, content):
    port = 465 #For SSL
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(user, password)
        server.sendmail(user, receiver, content)

def buildEmail(user, receiver, initialDate, finalDate, filePath):
    locale.setlocale(locale.LC_TIME, 'es-co')
    email = MIMEMultipart()

    if(initialDate.month == finalDate.month):
        dateText = "{} al {} de {}".format(initialDate.day, finalDate.day, initialDate.strftime("%B"))
        subject = "Factura {} {} - {} {}".format(initialDate.strftime("%B").capitalize(), initialDate.day, finalDate.day, initialDate.year)
    else:
        dateText = "{} de {} al {} de {}".format(initialDate.day, initialDate.strftime("%B"), finalDate.day, finalDate.strftime("%B"))
        subject = "Factura {} {} - {} {} {}".format(
            initialDate.strftime("%B").capitalize(),
            initialDate.day,
            finalDate.strftime("%B").capitalize(),
            finalDate.day, initialDate.year)

    body = """Hola Alex,

Te envío la factura correspondiente a la semana del {}.

Quedo atento a tus comentarios.

Muchas gracias. (Mensaje autogenerado por el invoiceGenerator :D)""".format(dateText)

    email['From'] = user

    email['Subject'] = subject
    
    email['To'] = receiver

    email.attach(MIMEText(body, "plain"))

    with open(filePath, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase(
            "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(attachment.read())


    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(filePath)}",
    )

    # Add attachment to message and convert message to string
    email.attach(part)
    text = email.as_string()

    # print(email)
    return text

def checkIfLoggedIn(session):
    return 'security_token' in session.cookies

def getTimezone(form):
    # timezoneInput = form.find(attrs = {"name": "user[time_zone]"})
    # print("Timezone input", timezoneInput)
    # return timezoneInput['value']
    return '-18000'

if __name__ == "__main__":
    main()
