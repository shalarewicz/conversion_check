import pandas as pd
from datetime import datetime
import sys

if sys.version_info[0] < 3:
    raise Exception("Python 3 or a more recent version is required.")

TEST_FILE = "Data Conversion Model.xlsx"
PROJECT_INFO = "Project Information"
HEADER_R0W = 5
VALID_VALUES = "Valid Values"

RA = "Research Administration"
RDC = "R&D"
IRB = "IRB"
IACUC = "IACUC"
SRS = "Research Safety"
IBC = "Biosafety"
DET = "Determinations"

ROW_OFFSET = 7

SUB_TYPES = ["AEO","MOD","CLS","REN","FUN","NEW","SBO","ORE","PDV","PUB","RES","REV","UPS"]
IACUC_SUB_TYPES = ["MOD","CLS","REN","FIP","FUN","NEW","SBO","PUB","RES","REV","RPE"]
REVIEW_TYPES = ["A","Q","M","E","C","F","L"]
IACUC_REVIEW_TYPES = ["A","Q","M","E","C","F","L","D"]
ACTIONS = ["ACK","APC","APP","CLS","DEF","EXE","FWD","INF",
           "NAP","NRE","NHR","RFB","SMR","SUS","TBL","TER","WDN"]
RD_ACTIONS = ["ACK","APC","APP","CLS","DEF","EXE","FWD",
              "INF","NAP","NRE","NHR","RFB","SMR","SUS","TBL","TER","WDN","CAP"]

RISK_LEVELS = ['MMR', 'MIN']

PROJECT_STATUSES = ["ACC","ACD","ALC","ACL","ACO","ACT","CLE","CLP","CLS","DIR",
                    "DIS","DMR","EMU","EXE","NRE","RNE","NHR","SUS","TER","WDN"]


def print_error(line, message, offset=ROW_OFFSET):
    print("ERROR: LINE", line + offset, ":", message)


def validate_date(date):
    try:
        # Excel gives timestamp as well
        datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        return True
    except ValueError:
        return False


def validate_encoding(text, encoding, line):
    try:
        text.encode(encoding, "strict")
    except UnicodeEncodeError as uee:
        print_error(line, "{enc} Special Character detected: '{c}' in {txt}".format(enc=encoding,
                                                                                    c=row[0][uee.start:uee.end],
                                                                                    txt=text))


def validate_text(text, encoding, line):
    validate_encoding(text, encoding, line)
    # Check for line breaks or special characters
    if any(c in text for c in ['\n', '\r', '\t', '\f', '\v', '\b', '\a']):
        print_error(line, "Special Character (Line Break) detected in {txt}".format(txt=text))


filename = TEST_FILE
try:
    filename = sys.argv[1]
except IndexError:
    if input("No file provided. Run with test file? Y/N") is 'Y':
        print("Running with test")
        filename = TEST_FILE
    else:
        sys.exit()

with pd.ExcelFile(filename) as xlsx:
    sheets = xlsx.sheet_names
    print("Detected Sheets", sheets)

    # Validate Project Information required fields
    print("\nValidating", PROJECT_INFO)
    project_info = pd.read_excel(xlsx, PROJECT_INFO, header=HEADER_R0W)
    for index, row in project_info.iterrows():
        # Check all cells for unsupported characters
        [(validate_text(str(col), 'latin-1', index) if pd.notnull(col) else "") for col in row]

        if pd.isnull(row[0]):
            print_error(index, "PROJECT TITLE required")
        if pd.isnull(row[1]):
            print_error(index, 'PI FIRST NAME required')
        if pd.isnull(row[2]):
            print_error(index, 'PI LAST NAME required')

    # Validate Review tabs
    sheets.remove(PROJECT_INFO)
    sheets.remove(VALID_VALUES)

    if IACUC not in sheets:
        print("No IACUC detected. Ensure tab is titled", IACUC)

    if RDC not in sheets:
        print("No R&DC detected. Ensure tab is titled", RDC)

    for sheet in sheets:
        print("\nValidating", sheet)
        reviews = pd.read_excel(xlsx, sheet, header=HEADER_R0W, usecols="B:O",
                                dtype={0:str, 5:str, 11:str, 12:str, 13:str})
        for index, row in reviews.iterrows():
            # Check all cells for unsupported characters
            [(validate_text(str(col), 'latin-1', index) if pd.notnull(col) else "") for col in row]

            # if any column contains a value, validate the row
            if row.notnull().values.any():
                # Check required fields
                # TODO Allow Pending Review print Warning
                # TODO Validate Vote fields print Warning if non-numeric

                if pd.isnull(row[0]):
                    print_error(index, "SUBMISSION DATE required")
                elif not validate_date(row[0]):
                    print_error(index, "SUBMISSION DATE {date} invalid".format(date=row[0]))

                if pd.isnull(row[1]):
                    print_error(index, 'SUBMISSION TYPE required')
                elif row[1] not in (IACUC_SUB_TYPES if sheet is IACUC else SUB_TYPES):
                    print_error(index, 'SUBMISSION TYPE {sub} invalid'.format(sub=row[1]))

                if pd.isnull(row[3]):
                    print_error(index, 'REVIEW TYPE required')
                elif row[3] not in (IACUC_REVIEW_TYPES if sheet is IACUC else REVIEW_TYPES):
                    print_error(index, 'REVIEW TYPE {rev} invalid'.format(rev=row[3]))

                if pd.isnull(row[4]):
                    print_error(index, 'ACTION required')
                elif row[4] not in (RD_ACTIONS if sheet is RDC else ACTIONS):
                    print_error(index, 'ACTION {act} invalid'.format(act=row[4]))

                if pd.isnull(row[5]):
                    print_error(index, 'EFFECTIVE DATE required')
                elif not validate_date(row[5]):
                    print_error(index, "EFFECTIVE DATE {date} invalid".format(date=row[5]))

                if pd.notnull(row[9]) and row[9] not in RISK_LEVELS:
                    print_error(index, 'RISK LEVEL {risk} invalid'.format(risk=row[10]))

                if pd.notnull(row[10]) and row[10] not in PROJECT_STATUSES:
                    print_error(index, 'PROJECT STATUS {status} invalid'.format(status=row[10]))

                if pd.notnull(row[11]) and not validate_date(row[11]):
                    print_error(index, "EXPIRATION DATE {date} invalid".format(date=row[11]))

                if pd.notnull(row[12]) and not validate_date(row[12]):
                    print_error(index, "INITIAL APPROVAL DATE {date} invalid".format(date=row[12]))

                if pd.notnull(row[13]) and not validate_date(row[13]):
                    print_error(index, "REPORT DUE {date} invalid".format(date=row[13]))

    print("Done!")
