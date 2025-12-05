"""Script to check students athletic eligibility based on grades in PowerSchool.

https://github.com/Philip-Greyson/D118-PS-Eligibility

Needs the google-api-python-client, google-auth-httplib2 and the google-auth-oauthlib:
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
also needs oracledb: pip install oracledb --upgrade
"""


import base64
import os
import mimetypes
from datetime import datetime
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

SCHOOL_CODES = ['5']
IGNORE_FULLYEAR_TERMS = False
NEEDED_GRADE_PERCENT = 60.0
NEEDED_PASSING_COURSES = 5
TO_EMAIL = os.environ.get('PS_ELIGIBILITY_EMAILS')  # emails that it will be sent to, can be mutliple emails comma separated
OUTPUT_FILE = 'whs_grades.txt'

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

        with open(OUTPUT_FILE, 'w') as output:
            today = datetime.now()
            # create the connecton to the PowerSchool database
            with oracledb.connect(user=DB_UN, password=DB_PW, dsn=DB_CS) as con:
                with con.cursor() as cur:  # start an entry cursor
                    for schoolCode in SCHOOL_CODES:
                        ineligibleStudents = {}  # for each school, create a list of ineligible students
                        outputText = ''  # for each school, make the final output string that is sent in the email
                        termlist = []  # make an empty list for valid terms per building
                        try:
                            termid = None
                            cur.execute("SELECT id, firstday, lastday, schoolid, dcid, isyearrec FROM terms WHERE schoolid = :school ORDER BY dcid DESC", school=schoolCode)  # get a list of terms for the school, filtering to not full years
                            terms = cur.fetchall()
                            for term in terms:  # go through every term
                                termStart = term[1]
                                termEnd = term[2]
                                isYear = term[5]
                                # compare todays date to the start and end dates
                                if ((termStart < today) and (termEnd > today)):
                                    termid = str(term[0])
                                    termDCID = str(term[4])
                                    if isYear == 1 and not IGNORE_FULLYEAR_TERMS:  # if the term we found is marked as a full year term and we arent ignoring them
                                        print(f'DBUG: Found yearlong term at building {schoolCode}: {termid} | {termDCID}')
                                        print(f'DBUG: Found yearlong term at building {schoolCode}: {termid} | {termDCID}', file=log)
                                        termlist.append(termid)  # add the term to the list we will iterate through. Comment out to ignore yearlong terms or set IGNORE_FULLYEAR_TERMS to True
                                    else:
                                        print(f'DBUG: Found good term at building {schoolCode}: {termid} | {termDCID}')
                                        print(f'DBUG: Found good term at building {schoolCode}: {termid} | {termDCID}', file=log)
                                        termlist.append(termid)  # add the term to the list we will iterate through. Comment out to ignore quarters, semesters, etc
                        except Exception as er:
                            print(f'ERROR while finding current term in building {schoolCode}: {er}')
                            print(f'ERROR while finding current term in building {schoolCode}: {er}', file=log)
                        try:
                            # find all courses with ath (athletics) in the name
                            cur.execute("SELECT c.course_number, c.course_name, c.dcid, ext.activity_start, ext.activity_end FROM courses c LEFT JOIN u_def_ext_courses ext ON c.dcid = ext.coursesdcid WHERE (instr(c.course_name, 'ATH-') > 0 OR instr(c.course_name, 'ACT-') > 0)")
                            courses = cur.fetchall()
                            for course in courses:
                                try:
                                    print(f'DBUG: Found course {course[1]}, checking to see if it is active')
                                    courseNum = course[0]
                                    courseName = course[1]
                                    courseDCID = course[2]
                                    activityStart = course[3]
                                    activityEnd = course[4]
                                    if activityStart and activityEnd:  # first check if there are even dates in the custom fields
                                        if ((activityStart < today) and (activityEnd > today)):
                                            # go through all the valid terms we found
                                            print(f'DBUG: Course {courseName} with DCID {courseDCID} is currently in season, continuing')
                                            print(f'DBUG: Course {courseName} with DCID {courseDCID} is currently in season, continuing', file=log)
                                            print(f'Students in athletics program {courseName} - {courseNum}:', file=output)
                                            for termid in termlist:
                                                print(f'DBUG: Looking for students in course {courseName} in term {termid} at building {schoolCode}')
                                                print(f'DBUG: Looking for students in course {courseName} in term {termid} at building {schoolCode}', file=log)
                                                try:
                                                    # next find all students in the current course in the current term if the activity is active
                                                    cur.execute('SELECT students.student_number, students.first_name, students.last_name, students.grade_level, students.id, cc.sectionid \
                                                                FROM cc LEFT JOIN students ON cc.studentid = students.id \
                                                                WHERE cc.course_number = :course AND cc.termid = :term AND cc.schoolid = :school', course=courseNum, term=termid, school=schoolCode)
                                                    students = cur.fetchall()
                                                    for student in students:
                                                        try:
                                                            print(f'DBUG: Found student {str(int(student[0]))} in course {courseName} during term ID {termid}')
                                                            print(f'DBUG: Found student {str(int(student[0]))} in course {courseName} during term ID {termid}', file=log)
                                                            studentEligible = False  # set the boolean flag to false each time to prevent rollover issues
                                                            numCoursesPassing = 0  # reset counter
                                                            failingClasses = []  # create an empty list that will keep track of what courses they are failing
                                                            passingClasses = []  # create an empty list of courses they are passing
                                                            stuNum = str(int(student[0]))
                                                            firstName = str(student[1]).title()
                                                            lastName = str(student[2]).title()
                                                            gradeLevel = int(student[3])
                                                            stuInternalID = str(int(student[4]))
                                                            try:
                                                                cur.execute('SELECT sg.finalgradename, sg.grade, sg.percent, sg.lastgradeupdate, sg.startdate, sg.enddate, courses.course_name FROM pgfinalgrades sg \
                                                                            LEFT JOIN sections ON sg.sectionid = sections.id \
                                                                            LEFT JOIN courses ON sections.course_number = courses.course_number \
                                                                            WHERE studentid = :stuid', stuid = stuInternalID)
                                                                grades = cur.fetchall()
                                                                for grade in grades:
                                                                    gradeTerm = str(grade[0])
                                                                    gradeLetter = str(grade[1])
                                                                    gradePercent = str(grade[2])
                                                                    gradeUpdateTime = grade[3].strftime('%m/%d/%Y - %H:%M:%S')
                                                                    classStart = grade[4].strftime('%m/%d/%Y')
                                                                    classEnd = grade[5].strftime('%m/%d/%Y')
                                                                    gradeCourseName = str(grade[6])
                                                                    if (grade[4] < datetime.now() < grade[5]):  # compare the start and end dates to the current date, only process if it contains the current date
                                                                        print(f'DBUG: {gradeCourseName}: {gradeLetter} {gradePercent} - from {classStart} to {classEnd} | Last updated {gradeUpdateTime}')
                                                                        if float(gradePercent) > NEEDED_GRADE_PERCENT:  # compare the grade percent to what we need to be considered passing
                                                                            passingClasses.append(f'{gradeCourseName} ({gradeTerm}): {gradePercent}% {gradeLetter} | Last updated {gradeUpdateTime}')  # append the course name, term and grade to the list
                                                                            numCoursesPassing+=1  # increment the counter for passing courses
                                                                        else:  # if they are not passing it, we want to keep track so we can print it out later
                                                                            failingClasses.append(f'{gradeCourseName} ({gradeTerm}): {gradePercent}% {gradeLetter} | Last updated {gradeUpdateTime}')  # append the course name, term and grade to the list
                                                                if numCoursesPassing >= NEEDED_PASSING_COURSES:
                                                                    print(f'DBUG: Student {stuNum} is eligible, passing {numCoursesPassing} classes')
                                                                    print(f'DBUG: Student {stuNum} is eligible, passing {numCoursesPassing} classes', file=log)
                                                                    print(f'\t•{stuNum} - {firstName} {lastName} is eligible and passing {numCoursesPassing} classes. Thse classes are the following:', file=output)
                                                                    for passedClass in passingClasses:
                                                                        print(f'\t\t•{passedClass}', file=output)
                                                                else:
                                                                    if not ineligibleStudents.get(stuNum, None):  # if they dont already exist in the dictionary
                                                                        ineligibleStudents.update({stuNum: failingClasses})  # add the student number
                                                                        print(f'INFO: Student {stuNum} in {courseName} is INELIGIBLE, they are only passing {numCoursesPassing} classes. They are failing {failingClasses}')
                                                                        print(f'INFO: Student {stuNum} in {courseName} is INELIGIBLE, they are only passing {numCoursesPassing} classes. They are failing {failingClasses}', file=log)
                                                                        if len(failingClasses) > 0:
                                                                            print(f'\t•{stuNum} - {firstName} {lastName} in {courseName} is INELIGIBLE because they are only passing {numCoursesPassing} classes. They are currently failing:',file=output)
                                                                            outputText += f'{stuNum} - {firstName} {lastName} in {courseName} is INELIGIBLE because they are only passing {numCoursesPassing} classes. They are currently failing:\n'  # add to the text that will be the body of the email
                                                                            for failedClass in failingClasses:
                                                                                print(f'\t\t•{failedClass}', file=output)
                                                                                outputText += f'\t•{failedClass}\n'  # add to the text that will be the body of the email
                                                                        else:
                                                                            print(f'\t•{stuNum} - {firstName} {lastName} in {courseName} is INELIGIBLE because they only have {numCoursesPassing} classes with grades. They are currently not failing any.', file=output)
                                                                            outputText += f'{stuNum} - {firstName} {lastName} in {courseName} is INELIGIBLE because they only have {numCoursesPassing} classes with grades. They are currently not failing any.\n'  # add to the text that will be the body of the email
                                                                    else:
                                                                        print(f'DBUG: Found duplicate student {stuNum} who is still ineligible')
                                                                        print(f'DBUG: Found duplicate student {stuNum} who is still ineligible', file=log)
                                                            except Exception as er:
                                                                print(f'ERROR while getting grades for {stuNum}: {er}')
                                                                print(f'ERROR while getting grades for {stuNum}: {er}', file=log)
                                                        except Exception as er:
                                                            print(f'ERROR while getting student info for {student[0]}: {er}')
                                                            print(f'ERROR while getting student info for {student[0]}: {er}', file=log)
                                                except Exception as er:
                                                    print(f'ERROR while getting students for course {courseName} in term {termid}: {er}')
                                                    print(f'ERROR while getting students for course {courseName} in term {termid}: {er}', file=log)
                                    else:
                                        print(f'WARN: Course {courseName} with DCID {courseDCID} has no dates entered for its start and end times')
                                        print(f'WARN: Course {courseName} with DCID {courseDCID} has no dates entered for its start and end times', file=log)
                                except Exception as er:
                                        print(f'ERROR while finding sections of course {courseName}: {er}')
                                        print(f'ERROR while finding sections of course {courseName}: {er}', file=log)
                        except Exception as er:
                            print(f'ERROR while finding courses with ATH in their names: {er}')
                            print(f'ERROR while finding courses with ATH in their names: {er}', file=log)
                    print(outputText)  # debug
        # should really put this inside the school and have different to emails, but right now only the high school cares so its fine      
        try:
            print(f'INFO: Sending email to {TO_EMAIL} with .csv file of grades')
            print(f'INFO: Sending email to {TO_EMAIL} with .csv file of grades', file=log)
            mime_message = EmailMessage()  # create an email message object
            # define headers
            mime_message['To'] = TO_EMAIL  # who the email gets sent to
            mime_message['Subject'] = f'WHS Eligibility - {datetime.now().strftime('%m/%d/%Y')}'  # subject line of the email
            mime_message.set_content(f'There are {len(ineligibleStudents.keys())} students that are ineligible. See the list below. Attached you should find a text file that contains a breakdown of all athletes sorted by which sport they are in, and how many courses they are passing. Please submit a ticket if there are issues or changes that need to be made.\n\n{outputText}')  # body of the email

            # add attachment of .csv output
            attachment_filename = OUTPUT_FILE
            # guessing the MIME type
            type_subtype, _ = mimetypes.guess_type(attachment_filename)
            maintype, subtype = type_subtype.split('/')

            with open(attachment_filename, 'rb') as fp:
                attachment_data = fp.read() # read the file data in and store it in the attachment_data
            mime_message.add_attachment(attachment_data, maintype, subtype, filename=f'{datetime.now().strftime('%m/%d/%Y')}-{OUTPUT_FILE}') # add the attacment data to the message object, give it a filename that was our pdf file name

            # encoded message
            encoded_message = base64.urlsafe_b64encode(mime_message.as_bytes()).decode()
            create_message = {'raw': encoded_message}
            send_message = (service.users().messages().send(userId="me", body=create_message).execute())
            print(f'DBUG: Email sent, message ID: {send_message["id"]}') # print out resulting message Id
            print(f'DBUG: Email sent, message ID: {send_message["id"]}', file=log)
        except HttpError as er:   # catch Google API http errors, get the specific message and reason from them for better logging
            status = er.status_code
            details = er.error_details[0]  # error_details returns a list with a dict inside of it, just strip it to the first dict
            print(f'ERROR {status} from Google API while sending eligibility email to {TO_EMAIL}: {details["message"]}. Reason: {details["reason"]}')
            print(f'ERROR {status} from Google API while sending eligibility email to {TO_EMAIL}: {details["message"]}. Reason: {details["reason"]}', file=log)
        except Exception as er:
            print(f'ERROR while sending eligibility email: {er}')
            print(f'ERROR while sending eligibility email: {er}', file=log)