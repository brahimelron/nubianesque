import os
import sys
import pyodbc
import datetime
import re
import argparse
import csv
import calendar

""" -------------------------------
  | I N I T I A L I Z A T I O N    |
    -------------------------------
"""
parser = argparse.ArgumentParser(description='Run a SQl query.',prefix_chars='-+/')
parser.add_argument('--pswd', action='store', dest='usrpwd',help='Enter your Windows password')
parser.add_argument('--jobcd', action='store', dest='jobcd',default='', \
                    help='Enter jobcode to process (MM=1278, RL=1279, MR=1277)')
parser.add_argument('--version', action='version', version='%(prog)s 1.3')
parser.add_argument('--uflg', action='store', dest='uflg', \
                    help='Enter 1 if you want the update the previous SQL template with current SQL.')

today = datetime.datetime.today()
one_day = datetime.timedelta(days=1)
thisyear = datetime.date.today().strftime("%Y")
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
priorMonth = lastMonth.strftime("%b")
# Calculating one year ago
one_year = datetime.timedelta(days=365)
yrfmt = "%Y"
last_year = today - one_year
prioryear = last_year.strftime(yrfmt)
usePriorYear_flg = 0
# Calculating therr days ago
three_day = datetime.timedelta(days=3)
dowkfmt = "%A"
format = "%Y-%m-%d"
if today.strftime(dowkfmt) == 'Monday':
    yesterday = today - three_day
else:
    yesterday = today - one_day
rptdate = yesterday.strftime(format)

userdomain = os.getenv('USERDOMAIN')
username = os.getenv('USERNAME')
userprof = os.getenv('USERPROFILE')

drvr = '{ODBC Driver 13 for SQL Server};'
tuser = userdomain + '\\' + username
tserver = 'PWEAMHANSQL2'
pserver = 'P6PO74SQL01'
tdatabase = 'InforProdTest'
pdatabase = 'InforProd'

# Get command line arguments
args = parser.parse_args()
tusrpwd = args.usrpwd
sqlqueryf = 'MES_M404_SummaryCnts.SQL'
inptPath = os.path.join(userprof, 'Documents', 'SQL', 'Monthly', sqlqueryf)

""" ------------------------------------
  | F U N C T I O N: put_back_prev_file  |
    ------------------------------------
"""
def put_back_prev_file(currPath):
    # --------------------------------------------
    # The function reverts the current file back  |
    # to previous version.                        |
    # --------------------------------------------
    prevfile = 'MES_M404_SummaryCnts(last).SQL'
    prevPath = os.path.join(userprof, 'Documents', 'SQL', 'Monthly', prevfile)
    if os.path.exists(prevPath):
        try:
            inptQueryF = open(prevPath, "r+")
            ouptQueryF = open(currPath, "w")
            for rl in inptQueryF.readlines():
                aline = str(rl)
                ouptQueryF.write(aline)
            inptQueryF.close()
            ouptQueryF.close()
            print('Previous version restored: ', prevfile)
            return 0
        except IOError:
            print ('File: ' + inptQueryF + ' cannnot be opened')
            return -1

""" ------------------------------------
  | F U N C T I O N: save_this_new_file  |
    ------------------------------------
"""
def save_this_new_file(currPath):
    # --------------------------------------------
    # The function reverts the current file back  |
    # to previous version.                        |
    # --------------------------------------------
    lastfile = 'MES_M404_SummaryCnts(last).SQL'
    lastPath = os.path.join(userprof, 'Documents', 'SQL', 'Monthly', lastfile)
    try:
        inptQueryF = open(currPath, "r+")
        ouptQueryF = open(lastPath, "w")
        for rl in inptQueryF.readlines():
            aline = str(rl)
            ouptQueryF.write(aline)
        inptQueryF.close()
        ouptQueryF.close()
        print('Last file version updated: ', lastfile)
        return 0
    except IOError:
        print ('File: ' + ouptQueryF + ' cannnot be opened')
        return -1

""" -----------------------------------
  | F U N C T I O N: set_jobcd_to_run  |
    -----------------------------------
"""
def set_jobcd_to_run(jobcode_arg):
    # ---------------------------------------------
    # The function builds the name of the SQL query |
    # file based on the jobcode argument if found.  |
    # ---------------------------------------------
    if jobcode_arg > ' ':
        if jobcode_arg == '1278':
            jbcd = 'MM'
        elif jobcode_arg == 'MM':
            jbcd = jobcode_arg
        elif jobcode_arg == '1277':
            jbcd = 'RL'
        elif jobcode_arg == 'RL':
            jbcd = jobcode_arg
        elif jobcode_arg == '1279':
            jbcd = 'MR'
        elif jobcode_arg == 'MR':
            jbcd = jobcode_arg
        else:
            jbcd = 'MM'
    else:
        jbcd = 'MM'
    print('Processing ' + jbcd + ' ' + jobcode_arg + '...')
    return jbcd

""" -----------------------------
  | F U N C T I O N: add_months  |
    -----------------------------
"""
def add_months(sourcedate,months):
    # ------------------------------------------------
    # The function calculates date increments in months |
    # ------------------------------------------------
    month = sourcedate.month - 1 + months
    year = int(sourcedate.year + month / 12 )
    month = month % 12 + 1
    day = min(sourcedate.day,calendar.monthrange(year,month)[1])
    return datetime.date(year,month,day)

""" -------------------------------
  | F U N C T I O N: put_rd_monthyr  |
    -------------------------------
"""
def put_rd_monthyr(filename):
    # ------------------------------------------------
    # The function determines the next month/year to   |
    # query and extract from the database and builds   |
    # the query to be run to pull that information.    |
    # ------------------------------------------------
    revisedf = 'newSQLQuery.SQL'
    revisedPath = os.path.join(userprof, 'Documents', 'SQL', 'Monthly', revisedf)
    if os.path.exists(revisedPath):
        os.remove(revisedPath)

    print("Prior Month is " + priorMonth.upper() + thisyear)
    if priorMonth.upper() == 'JAN':
        usePriorYear_flg = 1
    else:
        usePriorYear_flg = 0
    # --- begin get two months ago ---
    month_minus2 = add_months(today, -2)
    two_month_ago = month_minus2.strftime("%b")
    if usePriorYear_flg == 1:
        print('Two months ago is ' + two_month_ago.upper() + prioryear)
    else:
        print('Two months ago is ' + two_month_ago.upper() + thisyear)
    # --- end   get two months ago ---

    if usePriorYear_flg == 1:
        rdgPtrn = '\,\sSUM\(CASE\sWHEN\s\(YEAR\(CP\.ADDDTTM\)\s\=\s.*\s' + two_month_ago.upper() + prioryear
    else:
        rdgPtrn = '\,\sSUM\(CASE\sWHEN\s\(YEAR\(CP\.ADDDTTM\)\s\=\s.*\s' + two_month_ago.upper() + thisyear
    matchObj = re.compile(rdgPtrn)
    # --- begin parse the jobcode ---
    jbcPtrn = 'AND\sCP\.PROB\s=\s[0-9]{4}'
    matchJbc = re.compile(jbcPtrn)
    # --- end   parse the jobcode ---
    try:
        inptQueryF = open(filename, "r+")
        reviseQueryF = open(revisedPath, "w")
        for thisln in inptQueryF.readlines():
            m = matchObj.search(thisln)
            j = matchJbc.search(thisln)
            if m:
                monnmbr = lastMonth.strftime("%m")
                rpn = '\'' + str(monnmbr) + '\'' + ')'
                nxtlne = thisln[0:109] + rpn + ' THEN 1 ELSE 0 END' + ') AS ' + priorMonth.upper() + thisyear + '\n'
                reviseQueryF.write(thisln)
                reviseQueryF.write(nxtlne)
            elif j:
                if args.jobcd > ' ':
                    jobcode = args.jobcd
                    #print('Processing jobcode ' + jobcode)
                    nxtlne = re.sub(r'\d{4}', jobcode, thisln)
                    reviseQueryF.write(nxtlne)
                else:
                    reviseQueryF.write(thisln)
            else:
                reviseQueryF.write(thisln)
        inptQueryF.close()
        reviseQueryF.close()
    except IOError:
        print ('File: ' + inptQueryF + ' cannnot be opened')
        exit()
    finally:
    # Replace the sql query file with the contents of the revised file
        inptQueryF = open(revisedPath, "r")
        ouptQueryF = open(filename, "w")
        for rl in inptQueryF.readlines():
            aline = str(rl)
            ouptQueryF.write(aline)
        inptQueryF.close()
        ouptQueryF.close()

""" -------------------------------
  | F U N C T I O N: run_sql_file  |
    -------------------------------
"""
def run_sql_file(filename, connection):
    # ------------------------------------------------
    # The function takes a filename and a connection   |
    # as input and will run the SQL query on the given |
    # connection                                       |
    # ------------------------------------------------
    try:
        inptQueryF = open(filename)
    except:
        print ('File: ' + sqlqueryf + ' cannnot be opened')
        exit()
    sql = inptQueryF.read()
    inptQueryF.close()
    dbCursor = connection.cursor()
    dbCursor.execute(sql)
    row = dbCursor.fetchone()
    columns = [column[0] for column in dbCursor.description]
    try:
        destPath = os.path.join('//10.20.130.100', 'ibm-reports', 'MSOS', 'Monthly_MM_RL_MR_SummaryCnts')
        #  begin  D E B U G
        #print('Passing jobcode ' + args.jobcd + ' to set_jobcd_to_run')
        #  end    D E B U G
        whchjc = set_jobcd_to_run(args.jobcd)
        OUT_FILE = destPath + "\\MES-M404 Monthly_" + whchjc + "_thru_" + rptdate + ".txt"
        ofile = open(OUT_FILE, "w")
        # Print the column headers
        otline = ','.join(columns) + '\n'
        ofile.write(otline)
        # csv writer preparation
        writer = csv.writer(ofile)
    except:
        print("Output file destination error:", sys.exc_info()[0])
        raise
    try:
        while row is not None:
            writer.writerow(row)
            row = dbCursor.fetchone()
    except pyodbc.DataError as de:
        print("Para: run_sql_file DataError: " + str(de))
    except TypeError as te:
        print("Para: run_sql_file TypeError: " + str(te))
    except pyodbc.InternalError as ite:
        print("Para: run_sql_file InternalError")
        print(ite)
    except pyodbc.OperationalError as ope:
        print("Para: run_sql_file OperationalError")
        print(ope)
    except pyodbc.NotSupportedError as nse:
        print("Para: run_sql_file NotSupportedError")
        print(nse)
    except pyodbc.ProgrammingError as pe:
        print("Para: run_sql_file ProgrammingError")
        print(pe)        
    try:
        ofile.close()
    except:
        print("Output file close error in run_sql_file:", sys.exc_info()[0])
        print('Program ending')
        raise

""" -------------------------------
  | F U N C T I O N: main          |
    -------------------------------
"""
def main():
    rc = put_back_prev_file(inptPath)
    if rc == -1:
        exit(rc)
    try:
        put_rd_monthyr(inptPath)
        conn = pyodbc.connect(DRIVER=drvr, SERVER=pserver, DATABASE=pdatabase, UID=tuser, PWD=tusrpwd, trusted_connection='yes')
        run_sql_file(inptPath, conn)
    except:
        print("Unexpected error in main:", sys.exc_info()[0])
        raise
    finally:
        conn.close()
        print('Done')
    if args.uflg == 1:
        rc = save_this_new_file(inptPath)
        if rc == -1:
            exit(rc)
        elif rc == 0:
            print ('Ready for next run')
            exit(rc)
    exit()

""" -------------------------------
  | M A I N   P R O G R A M        |
    -------------------------------
"""
if __name__ == "__main__":
    main()
