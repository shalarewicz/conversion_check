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
BOARDS = [RA, RDC, IRB, IACUC, SRS, IBC, DET]

ROW_OFFSET = 7
VALID_VALUE_CATS = ["REV", "ACT", "PRJ", "RIS", "SUB"]
REVIEW_VV_CAT = "REV"
ACTION_VV_CAT = "ACT"
PRJ_STAT_VV_CAT = "PRJ"
RISK_VV_CAT = "RIS"
SUB_VV_CAT = "SUB"

SUB_TYPES = ["AEO","MOD","CLS","REN","FUN","NEW","SBO","ORE","PDV","PUB","RES","REV","UPS"]
IACUC_SUB_TYPES = ["MOD","CLS","REN","FIP","FUN","NEW","SBO","PUB","RES","REV","RPE"]

REVIEW_TYPES = ["A","Q","M","E","C","F","L"]

ACTIONS = ["ACK","APC","APP","CLS","DEF","EXE","FWD","INF","NAP","NRE","NHR","RFB","SMR","SUS","TBL","TER","WDN"]

RISK_LEVELS = ['MMR', 'MIN']

PROJECT_STATUSES = ["ACC","ACD","ALC","ACL","ACO","ACT","CLE","CLP","CLS","DIR",
                    "DIS","DMR","EMU","EXE","NRE","RNE","NHR","SUS","TER","WDN"]

MIN_DATE = datetime(1970,1,1)

MAX_DATE = datetime(2037,12,31)


def print_error(line, message, offset=ROW_OFFSET):
    print("ERROR: LINE", line + offset, ":", message)


def print_warning(line, message, offset=ROW_OFFSET):
    print("WARNING: LINE", line + offset, ":", message)


def validate_date(date):
    try:
        # pandas exports all date formats as YYYY-MM-DD HH:MM:SS
        # Unknown if visual format in Excel effects the conversion process. Support should inspect data visually
        d = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        return d >= MIN_DATE and d <= MAX_DATE

    except ValueError:
        return False


def validate_encoding(text, encoding, line):
    try:
        text.encode(encoding, "strict")
    except UnicodeEncodeError as uee:
        try:
            print_warning(line, "{enc} Special Character detected: '{c}' in {txt}".format(enc=encoding,
                                                                                   c=text[uee.start:uee.end],
                                                                                      txt=text))
        except UnicodeEncodeError as uee:
            print_error(line, "Special character detected. Could not print warning")



def validate_text(text, encoding, line):
    validate_encoding(text, encoding, line)
    # Check for line breaks or special characters
    if any(c in text for c in ['\n', '\r', '\t', '\f', '\v', '\b', '\a']):
        print_error(line, "Special Character (Line Break) detected in {txt}".format(txt=text))


def read_valid_values(filename):
    # Initialize the dict
    result = {}
    for board in BOARDS:
        inner = {}
        for cat in VALID_VALUE_CATS:
            inner[cat] = []
        result[board]=inner

    with pd.ExcelFile(filename) as xlsx:
        values = pd.read_excel(xlsx, "values",header=0,names=["cat", "val", "code", RA, DET, IRB, IACUC, IBC, SRS, RDC])
        for index, valid_value in values.iterrows():
            for board in BOARDS:
                if pd.notnull(valid_value[board]):
                    result.get(board).get(valid_value["cat"]).append(valid_value["code"].strip())

    return result


valid_values = read_valid_values("valid_values_map.xlsx")

filename = TEST_FILE
if len(sys.argv) > 1:
    filename = sys.argv[1]
else:
    if input("No file provided. Run with test file? Y/N") == 'Y':
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

        if ',' in str(row[6]):
            print_error(index, 'INTERNAL REFERENCE NUMBER cannot contain a comma')

    # Validate Review tabs
    sheets.remove(PROJECT_INFO)
    sheets.remove(VALID_VALUES)
    for sheet in sheets:
        if sheet not in BOARDS: 
            print("Invalid sheet name " + sheet + ". Ensure review sheets titled one of " + str(BOARDS))
            sys.exit();
    
    # Check for missing sheets
    # [print("No", board, "detected. Ensure tab is titled", board) if board not in sheets else "" for board in BOARDS]

    for sheet in sheets:
        print("\nValidating", sheet)
        reviews = pd.read_excel(xlsx, sheet, header=HEADER_R0W, usecols="B:O",
                                dtype={0: str, 5: str, 11: str, 12: str, 13: str})
        for index, row in reviews.iterrows():
            try:
                # Check all cells for unsupported characters
                [(validate_text(str(col), 'latin-1', index) if pd.notnull(col) else "") for col in row]

                # if any column contains a value, validate the row
                if row.notnull().values.any():
                    # Check required fields
                    if pd.isnull(row[0]):
                        print_error(index, "SUBMISSION DATE required")
                    elif not validate_date(row[0]):
                        print_error(index, "SUBMISSION DATE {date} invalid".format(date=row[0]))

                    if pd.isnull(row[1]):
                        print_error(index, 'SUBMISSION TYPE required')
                    elif row[1].strip() not in valid_values[sheet][SUB_VV_CAT]:
                        if row[1].strip() not in (IACUC_SUB_TYPES if sheet == IACUC else SUB_TYPES):
                            print_error(index, 'SUBMISSION TYPE {sub} invalid'.format(sub=row[1]))
                        else:
                            print_warning(index,
                                          'SUBMISSION TYPE {sub} not supported by board type but does not cause failure'
                                          .format(sub=row[1]))

                    # Allow Submissions Pending Review
                    pending = False
                    if pd.notnull(row[0]) and pd.notnull(row[1]) and pd.isnull(row[3]):
                        if pd.isnull(row[4]):
                            if pd.isnull(row[5]):
                                pending = True
                                # Submission is Pending Review check all subsequent columns are blank.
                                if any([pd.notnull(col) for col in row[6:]]):
                                    print_error(index, 'PENDING REVIEW but unvalidated review information entered')
                                else:
                                    print_warning(index, 'Line is PENDING REVIEW')

                    if not pending:
                        # Check Review Types
                        if pd.isnull(row[3]):
                            print_error(index, 'REVIEW TYPE required')
                        elif row[3].strip() not in valid_values[sheet][REVIEW_VV_CAT]:
                            if row[3].strip() not in REVIEW_TYPES:
                                print_error(index, 'REVIEW TYPE {rev} invalid'.format(rev=row[3]))
                            else:
                                print_warning(index,
                                              'REVIEW TYPE {rev} not supported by board type but does not cause failure'
                                              .format(rev=row[3]))

                        # Check Action
                        if pd.isnull(row[4]):
                            print_error(index, 'ACTION required')
                        elif row[4].strip() not in valid_values[sheet][ACTION_VV_CAT]:
                            if row[4].strip() not in ACTIONS:
                                print_error(index, 'ACTION {act} invalid'.format(act=row[4]))
                            else:
                                print_warning(
                                    index,
                                    'ACTION {act} not supported by board type but does not cause failure'
                                    .format(act=row[4]))

                        # Check Effective Date
                        if pd.isnull(row[5]):
                            print_error(index, 'EFFECTIVE DATE required')
                        elif not validate_date(row[5]):
                            print_error(index, "EFFECTIVE DATE {date} invalid".format(date=row[5]))

                        # Check for non-numeric votes
                        if any([pd.notnull(col) and not isinstance(col, int) and not isinstance(col, float) for col in row[6:8]]):
                            print_warning(index, "VOTE non numeric")

                        # Check Risk Level
                        if pd.notnull(row[9]) and row[9].strip() not in valid_values[sheet][RISK_VV_CAT]:
                            # SRS, IBC, RDC do not record risk level. Providing a risk level for these boards does not cause
                            # a failure. Don't reject a valid value
                            if row[9].strip() not in RISK_LEVELS or sheet not in [SRS, IBC, RDC, IACUC]:
                                print_error(index, 'RISK LEVEL {risk} invalid'.format(risk=row[9]))

                        # Check Project Status
                        if pd.notnull(row[10]) and row[10].strip() not in valid_values[sheet][PRJ_STAT_VV_CAT]:
                            if row[10].strip() not in PROJECT_STATUSES:
                                print_error(index, 'PROJECT STATUS {status} invalid'.format(status=row[10]))
                            else:
                                print_warning(index,
                                              'PROJECT STATUS {status} not supported by board but does not cause failure'
                                              .format(status=row[10]))

                        # Check Expiration Date
                        if pd.notnull(row[11]) and not validate_date(row[11]):
                            print_error(index, "EXPIRATION DATE {date} invalid".format(date=row[11]))

                        # Check Initial Approval Date
                        if pd.notnull(row[12]) and not validate_date(row[12]):
                            print_error(index, "INITIAL APPROVAL DATE {date} invalid".format(date=row[12]))

                        # Check Report Due Date
                        if pd.notnull(row[13]):
                            if not validate_date(row[13]):
                                print_error(index, "REPORT DUE {date} invalid".format(date=row[13]))

                            if row[13] == row[11]:
                                print_error(index,
                                            "REPORT DUE {date} cannot equal EXPIRATION DATE {exp}".format(date=row[13],
                                                                                                          exp=row[11]))
            except:
                print_error(index, "Error parsing line. Please review.")
    print("Done!")
