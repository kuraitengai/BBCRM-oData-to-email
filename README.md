# BBCRM oData to email
Using Blackbaud CRM, connect to a query via oData and embed the results in the body of an email.

## Installation
O365 is available on pypi.org. Simply run `pip install O365` to install it.

Requirements: >= Python 3.4

Project dependencies installed by pip:
* requests
* requests-oauthlib
* beatifulsoup4
* stringcase
* python-dateutil
* tzlocal
* pytz

## Setup
To use Gmail, you simply need access to a username and password.

To use Outlook, you simply need the program installed on the computer/server and have the program open.

To use OutlookOWA, you need to Authenticate and Authorize access. see https://github.com/O365/python-o365#authentication

## Usage
Update the odata_queries file for the specific output of your query.

Drop any additional fields not needed. (i.e. using df.drop for the QUERYRECID and any currencyID fields)

Reorder the query fields to the desired order since BBCRM likes to put them alphabetically sometimes.

Convert any date/time and currency fields from text (JSON output) to date/time/currency.

```python
    else:
        df = pd.json_normalize(r_json['value'])
        df = df.drop(['QUERYRECID'], axis=1)
        df = df[['LookupID','Name','ProspectManager','StartDate','SpouseManager','SpouseStartDate']]
        df = df.sort_values(['ProspectManager','StartDate'])
        df['StartDate'] = pd.to_datetime(df['StartDate'])
        df['SpouseStartDate'] = pd.to_datetime(df['SpouseStartDate'])
        df['StartDate'] = df['StartDate'].dt.date
        df['SpouseStartDate'] = df['SpouseStartDate'].dt.date
```

The utilities file contains code for sending emails via Gmail, Outlook, OutlookOWA with and without attachments). In the odata__assignmenterrors file, ONLY use the function for the ONE email service you actually want to use.
