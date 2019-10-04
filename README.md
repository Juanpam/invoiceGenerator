# invoiceGenerator
A python program to facilitate the generation of invoices, using a Web scraper


## Building

To prepare the program for execution just run the following command:

```
pip install -r requirements.txt
```

**Note:** This program is built using python 3!

## Running

To run the program you should edit the `config.json` file and change the properties to adjust to your use case. After that, just run the following command:

```
py main.py
```

on the project's root folder.

### Config File
```json=
{
    "username": "username",                // Assembla's username
    "password": "secret",                  //Assembla's password
    "name": "Rick Sanchez",                //Name for the invoice
    "path": ".",                           //Path to store the invoice file
    "initialDate": "01/01/2019",           //Initial date
    "finalDate": "02/01/2019",             //Final date
    "sEmail": "Y",                         //Do you want to send an email?
    "emailAddress": "youremail@gmail.com", //Sender email address (must be gmail)
    "emailPass": "secret",                 //Sender email password
    "receiver": "server@server.com"        //Recipient email address
}
```

### Notes
To send an email successfully, you have to permit access to insecure apps on Gmail. This way, the gmail smpt server will log you in correctly and send the email.

To change this setting, click [here](https://myaccount.google.com/lesssecureapps)