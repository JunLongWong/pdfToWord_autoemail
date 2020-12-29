# pdfToWord_autoemail
Automation project

This script extracts text from a pdf by locating pre-defined coordinates and populate data in a word document in a pre-defined format & let the user decide whether to send out an email with the attached word document output to the other party.

file - "senderParticulars" is not provided as it contains sensitive information.

file - "sensitiveInfomation" is not provided as it contains sensitive information.

## Specifications
Python 3.9.0

pdfminer library

## Prerequisition
### 1. Python
Please install python from the following [link](https://www.python.org/ftp/python/3.9.0/python-3.9.0-amd64.exe)

### 2. Python pip 
Pip should come packaged together with python installation.

To check if pip is installed, run "pip --version".

If in case of missing pip package:
1. Download from [link](https://bootstrap.pypa.io/get-pip.py).
2. Start command prompt or bash and "cd" to get-pip.py location.
3. Run python get-pip.py

### 3. Installation of python-docx
Please run following command to download pdfminer package.
1. pip install pdfminer

## Run Script
1. Open "cmd"
2. "cd" to project root directory 
3. Run "python manage.py pdf.py [pdfFilename]

