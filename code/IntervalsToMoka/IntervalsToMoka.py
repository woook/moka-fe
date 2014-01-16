"""
Prerequisites & installation:
    python 2.7.5 (python-2.7.5.msi)
    pyodbc 3.0.7 (pyodbc-3.0.7.win32-py2.7.exe)
    pandas 12 (pandas-0.12.0.win32-py2.7.exe)
    numpty 1.8.0 (numpy-1.8.0-win32-superpack-python2.7.exe)
    dateutil 1.5 (copy the dateutil folder into the python directory)

    All packages are in F:\Moka\Files\Software\IntervalsToMoka\
    Install python into C:\Array\Software\python27
    Then install pyodbc, pandas and numpty
    Then copy the dateutil folder
    F:\Moka\Files\Software\IntervalsToMoka\python-dateutil-1.5\dateutil\
    into the python27 directory

Script location:
    F:\Moka\Files\Analysis\IntervalsToMoka

Pyhton version:
    This script was created using python 2.7.5

Usage:
    IntervalsToMoka RunNumber
    This will search for an input file called RunNumber + _IntervalBasedReport.xls
    in F:\Moka\Files\Analysis\Year\Run\
    This will output a file called RunNumber + _IFM.csv
    in F:\Moka\Files\Analysis\Year\Run\
"""

import pandas as pd
import csv
import cStringIO
import codecs
import re
import pyodbc
import pandas.io.sql
import pandas as pd
import sys

class UnicodeWriter:
#A CSV writer which will write rows to CSV file "f", which is encoded in the given encoding.
    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        # Redirect output to a queue
        self.queue = cStringIO.StringIO()
        self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
        self.stream = f
        self.encoder = codecs.getincrementalencoder(encoding)()

    def writerow(self, row):
        self.writer.writerow([s.encode("utf-8") for s in row])
        # Fetch UTF-8 output from the queue ...
        data = self.queue.getvalue()
        data = data.decode("utf-8")
        # ... and reencode it into the target encoding
        data = self.encoder.encode(data)
        # write to the target stream
        self.stream.write(data)
        # empty queue
        self.queue.truncate(0)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)

#Make connection to database on the GSTTV Moka server
#for testing on devdatabase cnxn = pyodbc.connect("DRIVER={SQL Server};
#SERVER=GSTTV-MOKA; DATABASE=devdatabase;")
cnxn = pyodbc.connect("DRIVER={SQL Server}; SERVER=GSTTV-MOKA; DATABASE=mokadata;")
cursor = cnxn.cursor()

#pull chromosome lookup table and DNALabellingID-HybID-Cy3PatientID-Cy5PatientID
#table from moka and write to local csv files
SQL = 'SELECT * FROM dbo.v_Chromosome'
SQLHybPatient = 'SELECT * FROM dbo.v_DNALabellingHybPatients'

rows = cursor.execute(SQL).fetchall()
rowsHyb = cursor.execute(SQLHybPatient).fetchall()

with open("F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\temp\\chromosomelookup.csv", 'wb') as f:
    rows = [[unicode(x) for x in row] for row in rows]
    writer= UnicodeWriter(f)
    writer.writerow(["ChrID", "Chr"])
    writer.writerows(rows)
with open("F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\temp\\DNALabellingHybPatient.csv", 'wb') as f:
    rowsHyb = [[unicode(x) for x in row] for row in rowsHyb]
    writer= UnicodeWriter(f)
    writer.writerow(["DNALabellingID", "HybID", "Cy3InternalPatientID", "Cy5InternalPatientID"])
    writer.writerows(rowsHyb)

cnxn.commit()

cursor.close

#Generate a data frame where the headers for the columns start on line 17
#argv[1] is the name of the interval based report file which is supplied as an argument when calling the script
#test version: df = pd.read_table('F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\IntervalBasedReports\\130220_IntervalBasedReport.xls', header= 17)
##df = pd.read_table('F:\\Moka\\Files\\Analysis\\20' + sys.argv[1][0] + sys.argv[1][1] + '\\' + sys.argv[1] + '\\' + sys.argv[1] + '_IntervalBasedReport.xls', header= 17)
df = pd.read_table('F:\\Moka\\Files\\Analysis\\20' + sys.argv[1][0] + sys.argv[1][1] + '\\' + sys.argv[1] + '\\' + sys.argv[1] + '_IntervalBasedReport.xls', header= 17, names=['AberrationNo','Chr','Cytoband','Start','Stop','Probes','Amplification','Deletion','pval','Gene Names'])
chrom = pd.read_csv('F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\temp\\chromosomelookup.csv', header= 0)
#print df[:5]
#print df[:9]

#form a dictionary from the chromosome lookup table relating the Chr and ChrID columns to one another
group = chrom.groupby('Chr')

#Generate a list of Chr values from the group dictionary
chromdict = group['ChrID'].tolist()

"print chromdict"

#Generating the AmpDel column which is the sum of the Amplification and Deletion columns
df["AmpDel"] = (df.Amplification + df.Deletion)

#Remove the white space - white space from the Cytoband column
df["Cytoband"]= df["Cytoband"].str.replace('\s-\s', '')

#Generate a new dataframe called df2 from df but removing the Amplification, Deletion and pval columns
df2 = df.drop(['Amplification', 'Deletion','pval'], 1)

#Match anywhere where the aberration number appears and put it into the list 'match'. This generates a Boolean list indicating if there is a match or not to the pattern described
match = df2.AberrationNo.str.contains('\d{3}.\d\s\(\d{6}')

"print match[:9]"
#Initialise all values iun the for loop
truecounter = 0
rowcount = 0
AbNo=[]

#Generate AbNo column which has all rows filled with the correct Aberration No
for row in match:
        if row == True:
            truecounter = (rowcount - truecounter) + truecounter
            "print df2.AberrationNo[truecounter]"
            AbNo.append(df2.AberrationNo[truecounter])


        else:
            df2.AberrationNo[rowcount] = df2.AberrationNo[truecounter]
            "print df2.AberrationNo[truecounter]"
            AbNo.append(df2.AberrationNo[truecounter])

        rowcount += 1

"print AbNo"

#Add the column entitled AbNo to the df2 dataframe
df2["AbNO"] = AbNo

#drop all rows where there is an empty value in the Chr column and re-assign this to a new dataframe df3
df3 = df2.dropna(axis = 0, how = 'all', subset = ['Chr'])

#Remove AberrationNo column from the df3 dataframe
df3 = df3.drop("AberrationNo", 1)

#translate Chr number into correct chromosomal value
#initilise ChrNo list
ChrNo=[]
#Generate a new column ChrNo containing the correct Chromosome Number
for item in df3.Chr:
    ChrNo.append(repr(chromdict[item]))

print ChrNo

#strip the square brackets and L off the flanking sides of each string in ChrNo
#Initialise ChrNum column
ChrNum=[]
for item in ChrNo:
    ChrNum.append(item.strip('[L]'))

print ChrNum[1]

#Replace the values in ChrNo with the values in the list ChrNum
df3["ChrNo"] = ChrNum

#Generate a list of chromosome values without 'chr'
ChrNew=[]
for number in df3.Chr:
    ChrNew.append(number.strip('chr'))

print ChrNew

#Replace the values in Chr with the values in the list ChrNew
df3["Chr"] = ChrNew

print df3.ChrNo

#Generate a new column called Band which involves merging ChrNo and Cytoband together
df3["Band"] = (df3.Chr + df3.Cytoband)

#Write the df3 dataframe to file as a csv file
df3.to_csv(path_or_buf='F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\temp\\test1.csv', sep=',')

dfcy = pd.read_csv('F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\temp\\DNALabellingHybPatient.csv', header= 0)
df3 = pd.read_csv('F:\\Moka\\Files\\Analysis\\IntervalsToMoka\\temp\\test1.csv', header= 0)

InternalPatientID=[]
DNALabellingID=[]
Chr19=[]
Band19=[]
Start19=[]
Stop19=[]
Ratio=[]
Probes=[]

rowcount=0

#print dfcy.Cy3InternalPatientID[2]
#For each row of df3 iterate through each row of dfcy and if the Aberration Numbers match populate the following columns
for part in df3.AbNO:
    innerrow=0
    for row in dfcy.HybID:
        if row == part:
            InternalPatientID.extend([dfcy.Cy3InternalPatientID[innerrow], dfcy.Cy5InternalPatientID[innerrow]])
            DNALabellingID.extend([dfcy.DNALabellingID[innerrow]]*2)
            Ratio.extend([-df3.AmpDel[rowcount], df3.AmpDel[rowcount]])
            Chr19.extend([df3.ChrNo[rowcount]]*2)
            Band19.extend([df3.Band[rowcount]]*2)
            Start19.extend([df3.Start[rowcount]]*2)
            Stop19.extend([df3.Stop[rowcount]]*2)
            Probes.extend([df3.Probes[rowcount]]*2)

            innerrow += 1
        else:
            innerrow += 1
    rowcount += 1

#Transform column lists into column series
InternalPatientID = pd.Series(InternalPatientID)
DNALabellingID = pd.Series(DNALabellingID)
Ratio = pd.Series(Ratio)
Chr19 = pd.Series(Chr19)
Band19 = pd.Series(Band19)
Start19 = pd.Series(Start19)
Stop19 = pd.Series(Stop19)
Probes = pd.Series(Probes)

print InternalPatientID

#Generate a dataframe from the columns series generated in the previous step
##df4 = pd.DataFrame(zip(InternalPatientID, DNALabellingID, Ratio, Chr19, Band19, Start19, Stop19, Probes),  columns = ["InternalPatientID", "DNALabellingID", "Ratio", "Chr19", "Band19",
##"Start19", "Stop19"])
df4 = pd.DataFrame(zip(InternalPatientID, DNALabellingID, Ratio, Chr19, Band19, Start19, Stop19, Probes),  columns = ["InternalPatientID", "DNALabellingID", "Ratio", "Chr19", "Band19",
"Start19", "Stop19", "Probes"])

#argv[2] is the output file name to be supplied by user
df4.to_csv(path_or_buf='F:\\Moka\\Files\\Analysis\\20' + sys.argv[1][0] + sys.argv[1][1] + '\\' + sys.argv[1] + '\\' + sys.argv[1] + '_IFM.csv', sep=',')
