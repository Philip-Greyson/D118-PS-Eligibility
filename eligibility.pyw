"""Script to check students athletic eligibility based on grades in PowerSchool.

https://github.com/Philip-Greyson/D118-PS-Eligibility

Needs the google-api-python-client, google-auth-httplib2 and the google-auth-oauthlib:
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
also needs oracledb: pip install oracledb --upgrade
"""


import base64
import os
from datetime import datetime, timedelta
from email.message import EmailMessage

import oracledb  # needed for connection to PowerSchool server (oracle database)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# setup db connection
DB_UN = os.environ.get('POWERSCHOOL_READ_USER')  # username for read-only database user
DB_PW = os.environ.get('POWERSCHOOL_DB_PASSWORD')  # the password for the database account
DB_CS = os.environ.get('POWERSCHOOL_PROD_DB')  # the IP address, port, and database name to connect to
print(f'DBUG: Database Username: {DB_UN} |Password: {DB_PW} |Server: {DB_CS}')  # debug so we can see where oracle is trying to connect to/with

# Google API Scopes that will be used. If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

SCHOOL_ID = 5

if __name__ == '__main__':
    with open('eligibility_log.txt', 'w') as log:
        startTime = datetime.now()
        startTime = startTime.strftime('%H:%M:%S')
        print(f'INFO: Execution started at {startTime}')
        print(f'INFO: Execution started at {startTime}', file=log)
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        service = build('gmail', 'v1', credentials=creds)  # create the Google API service with just gmail functionality

        with open('whs_grades.csv', 'w') as output:
            print('Student Number,First Name,Last Name,Grade Level,Course Name,Term Name,Term Start,Term End,Grade Letter,Grade Percentage,Grade Points Earned,Grade Points Possible,Grade Last Updated', file=output)  # print header
            # create the connecton to the PowerSchool database
            with oracledb.connect(user=DB_UN, password=DB_PW, dsn=DB_CS) as con:
                with con.cursor() as cur:  # start an entry cursor
                    cur.execute('SELECT stu.student_number, stu.first_name, stu.last_name, stu.grade_level, courses.course_name, sg.finalgradename, sg.grade, sg.percent, sg.points, sg.pointspossible, sg.lastgradeupdate, sg.startdate, sg.enddate \
                                FROM PGFinalGrades sg LEFT JOIN students stu ON sg.studentid = stu.id \
                                LEFT JOIN sections ON sg.sectionid = sections.id \
                                LEFT JOIN courses ON sections.course_number = courses.course_number \
                                WHERE stu.schoolid = :school AND stu.enroll_status = 0 ORDER BY sg.studentid', school=SCHOOL_ID)
                    grades = cur.fetchall()
                    for entry in grades:
                        stuNum = str(int(entry[0]))
                        firstName = str(entry[1])
                        lastName = str(entry[2])
                        grade = int(entry[3])
                        course = str(entry[4])
                        termName = str(entry[5])
                        gradeLetter = str(entry[6])
                        gradePercent = str(entry[7])
                        gradePoints = int(entry[8])
                        gradePointsPossible = int(entry[9])
                        # gradeUpdateTime = entry[10]
                        # print(type(gradeUpdateTime))
                        gradeUpdateTime = entry[10].strftime('%H:%M:%S - %m/%d/%Y')
                        classStart = entry[11].strftime('%m/%d/%Y')
                        classEnd = entry[12].strftime('%m/%d/%Y')
                        if (entry[11] < datetime.now() < entry[12]):  # compare the start and end dates to the current date, only print out those that contain the current date
                            print(f'DBUG: {stuNum}-{firstName} {lastName} in grade {grade} has course {course} during term {termName} that goes from {classStart} to {classEnd} and current grade {gradeLetter} - {gradePercent} with {gradePoints} out of {gradePointsPossible}, last updated {gradeUpdateTime}')
                            print(f'DBUG: {stuNum}-{firstName} {lastName} in grade {grade} has course {course} during term {termName} that goes from {classStart} to {classEnd} and current grade {gradeLetter} - {gradePercent} with {gradePoints} out of {gradePointsPossible}, last updated {gradeUpdateTime}', file=log)
                            print(f'{stuNum},{firstName},{lastName},{grade},{course},{termName},{classStart},{classEnd},{gradeLetter},{gradePercent},{gradePoints},{gradePointsPossible},{gradeUpdateTime}', file=output)
                        # else:
                        #     print(f'WARN: Found grade for student {stuNum} in course {course} for term {termName} that goes from {classStart} to {classEnd} that is not the current term, skipping')
                        #     print(f'WARN: Found grade for student {stuNum} in course {course} for term {termName} that goes from {classStart} to {classEnd} that is not the current term, skipping', file=log)