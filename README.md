# OTRS Calendar to Excel

This Python script connects directly to the MySQL/MariaDB of an OTRS 
installation and exports a single calendar to an Excel spreadsheet. The
current, last and next year are included.

Only calendar appointments that have a resource assigned will be included. 
Assigning resources is a feature of the commercial OTRS Business Solution 
edition.


### Installation

System

```
apt-get install python-virtualenv
```

OR 

```
yum install python-virtualenv
```

Python

```
virtualenv -p python3 venv
source venv/bin/activate
pip install -U pip
pip install -r requirements.txt
```

### Config

Create a file called `config.py` and fill in your settings:

```
# MySQL Connection Settings
MYSQL_HOST = "localhost"
MYSQL_USER = "otrs"
MYSQL_PASS = "<insert_password_here>"
MYSQL_DB = "otrs"

CALENDER_ID = 1
APPOINTMENT_SYMBOL = "U"
``` 
