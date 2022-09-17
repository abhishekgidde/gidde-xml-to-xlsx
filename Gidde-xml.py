""" 
Made by Abhishek C. Gidde on demand for Ranjit Kadam
"""
from bs4 import BeautifulSoup as giddesoup
import pandas as pd
import argparse

parser = argparse.ArgumentParser()
parser.add_argument("-f", "--filename", dest = "filename", help="Name of the XML file", required=True)
parser.add_argument("-r", "--resultfile", dest = "resultfile",default="gidde.xlsx" ,help="Name of the .Xlsx file to be created")
parser.add_argument("-v", "--verbose", dest = "verbose",default=False ,help="Verbose mode: set -v True and You can see the output in the terminal")

args = parser.parse_args()

resultPath = "./Result/"+parser.parse_args().resultfile
xmlFile = parser.parse_args().filename
verbose = parser.parse_args().verbose

if resultPath.endswith(".xlsx") == False:
    resultPath = resultPath + ".xlsx"

if xmlFile.endswith(".xml") == False:
    xmlFile = xmlFile + ".xml"


dataVerbose = False

# check if verbose mode is enabled
if verbose == "true" or verbose == "True":
    dataVerbose = True


print("resultPath: ",resultPath)
print("xmlFile: ",xmlFile)


try:
    with open(xmlFile,'r') as f:
     data = f.read()
except:
    print("File not found: ",xmlFile)
    exit()

writer = pd.ExcelWriter(resultPath, engine = 'openpyxl')
wb = writer.book

bs_data = giddesoup(data,"xml")
tag_appVisibility = bs_data.find_all("applicationVisibilities")

#Tag Application Visibilities
print("----------- Application Visibilites ------------")
appNames = []
appDefaultBools = []
appVisibleBools = []
for single_tag in tag_appVisibility:
    for i in single_tag.find_all("application"):
        appNames.append(str(i.contents[0]))
    for j in single_tag.find_all("default"):
        appDefaultBools.append(str(j.contents[0]))
    for k in single_tag.find_all("visible"):
        appVisibleBools.append(str(k.contents[0]))

df = pd.DataFrame({'App Name':appNames,'App Default': appDefaultBools,'App Visible': appVisibleBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='appVisibility')


#Tag Class Accesses
print("----------- Class Accesses ------------")
apexClassNames = []
apexClassEnableBools = []

classAccesses = bs_data.find_all("classAccesses")
for single_tag in classAccesses:
    for i in single_tag.find_all("apexClass"):
        apexClassNames.append(str(i.contents[0]))
    for j in single_tag.find_all("enabled"):
        apexClassEnableBools.append(str(j.contents[0]))

df = pd.DataFrame({'Apex Class Name':apexClassNames,'Apex Class Enable': apexClassEnableBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='classAccesses')


#Tag Field Accesses
print("----------- Custom Meta Data Type Accesses------------")
customMetaDataTypeNames = []
customMetaDataTypeEnableBools = []
customMetaDataTypeAccesses = bs_data.find_all("customMetadataTypeAccesses")

for single_tag in customMetaDataTypeAccesses:
    for i in single_tag.find_all("name"):
        customMetaDataTypeNames.append(str(i.contents[0]))
    for j in single_tag.find_all("enabled"):
        customMetaDataTypeEnableBools.append(str(j.contents[0]))

df = pd.DataFrame({'Custom Meta Data Type Name':customMetaDataTypeNames,'Custom Meta Data Type Enable': customMetaDataTypeEnableBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='customMetaDataTypeAccesses')

#Tag externalDataSourceAccesses
print("----------- External Data Source Accesses ------------")
externalDataSourceNames = []
externalDataSourceEnableBools = []
externalDataSourceAccesses = bs_data.find_all("externalDataSourceAccesses")
for single_tag in externalDataSourceAccesses:
    for i in single_tag.find_all("externalDataSource"):
        externalDataSourceNames.append(str(i.contents[0]))
    for j in single_tag.find_all("enabled"):
        externalDataSourceEnableBools.append(str(j.contents[0]))
df = pd.DataFrame({'External Data Source Name':externalDataSourceNames,'External Data Source Enable': externalDataSourceEnableBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='externalDataSourceAccesses')


#Tag fieldPermissions
print("----------- Field Permissions ------------")
fieldPermissionsFieldNames = []
fieldPermissionsEditableBools = []
fieldPermissionsReadableBools = []
fieldPermissionsAccesses = bs_data.find_all("fieldPermissions")

for single_tag in fieldPermissionsAccesses:
    for i in single_tag.find_all("field"):
        fieldPermissionsFieldNames.append(str(i.contents[0]))
    for j in single_tag.find_all("editable"):
        fieldPermissionsEditableBools.append(str(j.contents[0]))
    for k in single_tag.find_all("readable"):
        fieldPermissionsReadableBools.append(str(k.contents[0]))

df = pd.DataFrame({'Field Permissions Field Name':fieldPermissionsFieldNames,'Field Permissions Editable': fieldPermissionsEditableBools,'Field Permissions Readable': fieldPermissionsReadableBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='fieldPermissions')

#Tag flowAccesses
print("----------- Flow Accesses ------------")
flowNames = []
flowEnableBools = []
flowAccesses = bs_data.find_all("flowAccesses")

for single_tag in flowAccesses:
    for i in single_tag.find_all("flow"):
        flowNames.append(str(i.contents[0]))
    for j in single_tag.find_all("enabled"):
        flowEnableBools.append(str(j.contents[0]))

df = pd.DataFrame({'Flow Name':flowNames,'Flow Enable': flowEnableBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='flowAccesses')


#Tag layoutAssignments
print("----------- Layout Assignments ------------")
layoutAssignmentsLayout = []
layoutAssignmentsRecordType = []

for single_tag in bs_data.find_all("layoutAssignments"):
    if(len(single_tag.find_all("recordType"))==0):
        layoutAssignmentsRecordType.append("NA")
    for i in single_tag.find_all("layout"):
        layoutAssignmentsLayout.append(str(i.contents[0]))
    for j in single_tag.find_all("recordType"):
        layoutAssignmentsRecordType.append(str(j.contents[0]))

df = pd.DataFrame({'Layout Assignments Layout':layoutAssignmentsLayout,'Layout Assignments Record Type': layoutAssignmentsRecordType})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='layoutAssignments')


#Tag objectPermissions
print("----------- Object Permissions ------------")
objectPermissionsObjectName = []
objectPermissionsAllowCreateBools = []
objectPermissionsAllowDeleteBools = []
objectPermissionsAllowEditBools = []
objectPermissionsAllowReadBools = []
objectPermissionsModifyAllRecordsBools = []
objectPermissionsViewAllRecordsBools = []

for single_tag in bs_data.find_all("objectPermissions"):
    for i in single_tag.find_all("object"):
        objectPermissionsObjectName.append(str(i.contents[0]))
    for j in single_tag.find_all("allowCreate"):
        objectPermissionsAllowCreateBools.append(str(j.contents[0]))
    for k in single_tag.find_all("allowDelete"):
        objectPermissionsAllowDeleteBools.append(str(k.contents[0]))
    for l in single_tag.find_all("allowEdit"):
        objectPermissionsAllowEditBools.append(str(l.contents[0]))
    for m in single_tag.find_all("allowRead"):
        objectPermissionsAllowReadBools.append(str(m.contents[0]))
    for n in single_tag.find_all("modifyAllRecords"):
        objectPermissionsModifyAllRecordsBools.append(str(n.contents[0]))
    for o in single_tag.find_all("viewAllRecords"):
        objectPermissionsViewAllRecordsBools.append(str(o.contents[0]))

df = pd.DataFrame({'Object Permissions Object Name':objectPermissionsObjectName,'Object Permissions Allow Create': objectPermissionsAllowCreateBools,'Object Permissions Allow Delete': objectPermissionsAllowDeleteBools,'Object Permissions Allow Edit': objectPermissionsAllowEditBools,'Object Permissions Allow Read': objectPermissionsAllowReadBools,'Object Permissions Modify All Records': objectPermissionsModifyAllRecordsBools,'Object Permissions View All Records': objectPermissionsViewAllRecordsBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='objectPermissions')



#Tag pageAccesses
print("----------- Page Accesses ------------")
pageNames = []
pageEnableBools = []

for single_tag in bs_data.find_all("pageAccesses"):
    for i in single_tag.find_all("apexPage"):
        pageNames.append(str(i.contents[0]))
    for j in single_tag.find_all("enabled"):
        pageEnableBools.append(str(j.contents[0]))

df = pd.DataFrame({'Page Name':pageNames,'Page Enable': pageEnableBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='pageAccesses')

#Tag recordTypeVisibilities
print("----------- Record Type Visibilities ------------")
recordTypeVisibilitiesRecordType = []
recordTypeVisibilitiesVisibleBools = []

for single_tag in bs_data.find_all("recordTypeVisibilities"):
    for i in single_tag.find_all("recordType"):
        recordTypeVisibilitiesRecordType.append(str(i.contents[0]))
    for j in single_tag.find_all("visible"):
        recordTypeVisibilitiesVisibleBools.append(str(j.contents[0]))

df = pd.DataFrame({'Record Type Visibilities Record Type':recordTypeVisibilitiesRecordType,'Record Type Visibilities Visible': recordTypeVisibilitiesVisibleBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='recordTypeVisibilities')



#Tag tabVisibilities
print("----------- Tab Visibilities ------------")
tabVisibilitiesTab = []
tabVisibilitiesVisibleBools = []

for single_tag in bs_data.find_all("tabVisibilities"):
    for i in single_tag.find_all("tab"):
        tabVisibilitiesTab.append(str(i.contents[0]))
    for j in single_tag.find_all("visibility"):
        tabVisibilitiesVisibleBools.append(str(j.contents[0]))

df = pd.DataFrame({'Tab Visibilities Tab':tabVisibilitiesTab,'Tab Visibilities Visible': tabVisibilitiesVisibleBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='tabVisibilities')


#Tag userPermissions
print("----------- User Permissions ------------")
userPermissionsName = []
userPermissionsEnabledBools = []

for single_tag in bs_data.find_all("userPermissions"):
    for i in single_tag.find_all("name"):
        userPermissionsName.append(str(i.contents[0]))
    for j in single_tag.find_all("enabled"):
        userPermissionsEnabledBools.append(str(j.contents[0]))

df = pd.DataFrame({'User Permissions Name':userPermissionsName,'User Permissions Enabled': userPermissionsEnabledBools})
print("Successful")
if(dataVerbose):print(df)
df.to_excel(writer, sheet_name='userPermissions')

writer.save()
writer.close()

