from openpyxl import Workbook

# Read directory of Excel files
xl_dir = '../MSD Holdings 2016/'
from os import listdir
from os.path import isfile, join
included_extensions = ['xlsx']
xl_files = [f for f in listdir(xl_dir) if any(f.endswith(ext) for ext in included_extensions)]
xl_files.sort()
#print xl_files[0:4] # TEMP\n"

#Index file, tab delimited
index_file_path = 'AAA_index.txt'
index_heading = 'FundID\tSecID\tFundName\tInceptionDate\tEndDate\tCSVDate\tNumHoldings\tFileName\n'

import os
import codecs
if os.path.exists(index_file_path):
    index_file = codecs.open(index_file_path, 'a',encoding='utf8')
else:
    index_file = codecs.open(index_file_path, 'w',encoding='utf8')
    index_file.write(index_heading)


# Outer loop
import datetime
overwrite = False # May use
csv_dir = 'csv/'
if not os.path.exists(csv_dir):
    os.makedirs(csv_dir)
#loop_from = 0
loop_to = 2

# Shrink list for testing # TEMP
xl_files = xl_files#[loop_from:loop_to]

for xl_file_name in xl_files:
    print "\n"
    print xl_file_name
    xl_file_full_path = xl_dir + xl_file_name
    from openpyxl import load_workbook
    wb = load_workbook(filename=xl_file_full_path, read_only=True) #data_only=True) ##
    sheet_names = wb.get_sheet_names()
    ##sheet_ranges = wb['range names'] 
    ##print(sheet_ranges['D18'].value)
    # Load Sheet1, local index
    ws = wb['Sheet1']
    r = 1
    fundIDs = [] 
    # fundID is index for the following dicts
    secIDs = {} 
    fundNames = {}
    inceptionDates = {}
    endDates = {}
    # Read local index sheet
    for row in ws.rows:
        ##for cell in row:
            ##print(cell.value)
        #if r == 1:
            #heading
            #print 'Heading: ' + row[0].value
        if r > 1:
            #content
            # Get fundID from columnB/row[1], etc
            # Check first column status before using
            if (row[0].value==u'Done!'):
                fundID = row[1].value
                fundIDs.append(fundID)
                secIDs[fundID]         = row[2].value
                fundNames[fundID]      = row[3].value
                inceptionDates[fundID] = row[4].value
                endDates[fundID]       = row[5].value
        r=r+1
    print fundIDs
    #print secIDs
    #print fundNames
    #print inceptionDates
    #print endDates
    # For each fundID, load sheet
    for fundID in fundIDs: 
        print fundID
        ws = wb[fundID]
        r = 1
        months = {} # reference dates by column index (from 3=D)
        # Create output file, tab delim
        csv_file_name = fundID + ".txt"
        csv_file = codecs.open(csv_dir + csv_file_name, 'w',encoding='utf8')
        #TODO overwrite check
        csv_file.write(u'holdingID\tholdingName\tholdingType\tdate\tvalue\n')
        for row in ws.rows:
            #dict of column index to month/date
            if r == 2:
                c = 3
                # Get dates from row 2, columnD/row[3] to end
                for column in row[3:]:
                    try:
                        months[c] = column.value.strftime("%Y%m%d")
                    except AttributeError:
                        months[c] = "" # OR should this be skipped?
                    c = c+1
                #print 'Fund ' + fundID + ' starts on '
                #print row[3].value
                #print u'Fund ' + fundID + ' starting ' + months[3]
            elif (r>2): #TEMP END EARLY and r < 5
                c = 3
                # Pull out details for one holding
                holdingID = row[0].value
                holdingName = row[1].value
                holdingType = row[2].value
                for column in row[3:]:
                    # Get the date of this column if it has content
                    if column.value is not None: #?  and len(column.value) > 0
                        if not holdingID:
                            holdingID = u''
                        try:
                            line = holdingID.encode('utf-8') 
                        except AttributeError:
                            line = str(holdingID)
                        try:
                            line = line + u'\t' + holdingName
                        except TypeError:
                            #print fundID
                            #print holdingID
                            #print len(line)
                            #print line
                            #print holdingName
                            line = line + u'\t' + str(holdingName)
                        line = line + u'\t' + holdingType
                        line = line + u'\t' + months[c]
                        line = line +  u'\t' + str(column.value) + u'\n'
                        csv_file.write(line)
                        #csv_file.write(holdingID + u'\t' + str(holdingName) + u'\t' + holdingType + u'\t' + str(months[c]) + u'\t' + str(column.value) + u'\n')
                    c = c+1
                #...
            r=r+1
        try: 
            csv_file.close()
        except IOError:
            print "IOError closing csv_file", csv_file_name
    # Write to index file
    #index_heading = 'FundID\tSecID\tFundName\tInceptionDate\tEndDate\tCSVDate\tNumHoldings\n'
    index_file.write(fundID + u'\t' + secIDs[fundID] + u'\t' + fundNames[fundID]+ u'\t' + inceptionDates[fundID].strftime("%Y%m%d")+ u'\t' + endDates[fundID].strftime("%Y%m%d")+ u'\t' + datetime.datetime.now().strftime("%Y%m%dT%H%M%S")+ u'\t' +str(r-2) + u'\t' + xl_file_name + u'\n')
    try: 
        index_file.flush()
    except IOError:
        print "IOError flushing index file"



# Wrap it up
try: 
    index_file.close()
except IOError: 
    True

# Reset while testing
#os.remove(index_file_path)
