# system imports
import logging
import urlparse

# app engine apis
from google.appengine.api import taskqueue

# flask
from flask import Flask, request, render_template
from werkzeug.exceptions import HTTPException, Aborter, default_exceptions
from werkzeug.http import HTTP_STATUS_CODES

# google api client library imports
from apiclient.discovery import build
from googleapiclient.errors import HttpError
from oauth2client.contrib.flask_util import UserOAuth2
import httplib2

# mailjet
import mailjet_rest
import requests_toolbelt.adapters.appengine

# local imports
import config
from utils.pdf import create_cert

# Use the App Engine requests adapter to allow the requests library to be
# used on App Engine.
requests_toolbelt.adapters.appengine.monkeypatch()

app = Flask(__name__)
app.config.from_object(config)
app.config['SECRET_KEY'] = 'random*strangageasf'
app.config['GOOGLE_OAUTH2_CLIENT_SECRETS_FILE'] = 'client_secret.json'
logger = logging.getLogger(__name__)

MAILJET_API_KEY = app.config['MJ_API_KEY']
MAILJET_API_SECRET = app.config['MJ_API_SECRET']
MAILJET_SENDER = app.config['MJ_SENDER']

oauth2 = UserOAuth2(
    app,
    scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])


################################################################################
# Helpers

class SheetNotFound(HttpError):
    """Throw this when sheet is not found"""
    pass


class EmailSendingUnavailable(HTTPException):
    """EmailSending unvailable due to overQuota. """
    code = 503
    description = ("Some emails couldn't be sent at the moment. "
                   "Please try again a bit later")

abort = Aborter()
default_exceptions[503] = EmailSendingUnavailable
HTTP_STATUS_CODES[503] = 'Email Sending Over Quota, Try again Later'


def get_sheet_id(url):
    path = urlparse.urlsplit(url).path
    try:
        sheet_id = path.split('/')[3]
    except IndexError:
        logger.exception('url passed is {}'.format(url))
        sheet_id = ''
    return sheet_id


def send_cert(name, email):
    """Send certificates via provided mail
    """
    certificate = create_cert(name)
    client = mailjet_rest.Client(
        auth=(MAILJET_API_KEY, MAILJET_API_SECRET))
    data = {
        'FromEmail': MAILJET_SENDER,
        'FromName': 'Wineloya Digital Advertising',
        'Subject': 'Subject: Digital Skills Training Certificate',
        'Text-part': (
            'Congratulations!, your digital skills training completion'
            'certificate is here!'),
        'Recipients': [{'Email': email}],
        'Attachments': [{
            "Content-type": "application/pdf",
            "Filename": "Certificate.pdf",
            "content": certificate
        }],
    }
    result = client.send.create(data=data)
    logger.info(result.json())


def sheets_client():
    """build connection to google sheets service
    """
    credentials = oauth2.credentials
    http = httplib2.Http()
    http = credentials.authorize(http)

    return build('sheets', 'v4', http=oauth2.http())


################################################################################
# Routes

@app.errorhandler(503)
def errormail(e):
    return e, 503


@app.route('/emailtask', methods=['POST'])
def email_task():
    values = request.values.getlist('values')
    import pdb
    pdb.set_trace()

    # for value in values:
    #     last_name = value[1]
    #     first_name = value[2]
    #     name = '{} {}'.format(last_name, first_name)
    #     email = value[4]
    #     send_cert(name, email)
    return ('OK')


@app.route('/sheet', methods=['GET', 'POST'])
@oauth2.required
def sheets():
    sheet_id = ''
    if request.method == 'POST':
        sheet_id = get_sheet_id(request.form['sheet'])
        service = sheets_client()
        logger.info('Spreadsheet with id: {} was accessed'.format(sheet_id))
        try:
            result = service.spreadsheets().values().get(
                spreadsheetId=sheet_id, range='A2:E').execute()
            values = result.get('values', [])

            taskqueue.add(url='/emailtask',
                          params={'values': values})
        except HttpError, err:
            if err.resp.status in [404]:
                # reason = json.loads(err.content).reason
                # return json.dumps(reaeon)
                msg = '{}: Check that the sheet provided is valid: {}'.format(
                    err.resp.reason, sheet_id)
                logger.error(msg)
                return msg
            else:
                raise
        # return json.dumps(values)
        return render_template('finish.html')
    return render_template('sheet.html')


@app.route('/single', methods=['GET', 'POST'])
def single():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        if email:
            send_cert(name, email)

        render_template('single.html')
    return render_template('single.html')


@app.route('/')
def home():
    return render_template("home.html")


if __name__ == '__main__':
    app.run(debug=True)
