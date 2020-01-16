import xlrd,xlwt,copy
import sys
import os
from datetime import date,datetime,time
print("----------------------------------------")
print(" ")
print("   Checkin Parser by  vincent@tech     ")
print("       use : python parser.py source.xlsx 2019-01-01 2019-10-11")
print(" ")
print("----------------------------------------")

if  len(sys.argv)  == 1 :
    exit(0)
print(sys.argv)
source = os.getcwd() + "/"+sys.argv[1]
resultFile = os.getcwd()+"/result.xls"

if len(sys.argv) >= 3 : 
    dateStart = datetime.fromisoformat(sys.argv[2] + " 00:00:00")
else :
    dateStart = None 

if len(sys.argv) >= 4 : 
    dateEnd = datetime.fromisoformat(sys.argv[3] + " 00:00:00")
else :
    dateEnd = None  

print("- parsing " + resultFile)


def dateInRange(t) : 
    if t == "" or t == None :
        return False
    t1 = datetime.fromisoformat(t)
    if dateStart != None : 
        if t1 < dateStart :
            return False
    
    if dateEnd != None : 
        if t1 > dateEnd :
            return False

    return True
 

# checkin records
class Record : 
    name = ""
    date = ""
    depart  = ""
    checkin  = ""
    checkout = "" 
    worktime = 0
    worktimeMin = 0 
    def __str__(self):
        return self.name
    def updateWorktime(self):
        if self.checkout != "" : 
            checkoutTime = datetime.fromisoformat(self.checkout)
            checkinTime  = datetime.fromisoformat(self.checkin)
            overworkTime = datetime.fromisoformat(checkoutTime.date().isoformat() + " 19:30:00")
            #Satday Sunday count all time  
            if(checkoutTime.weekday == 5 or checkoutTime.weekday == 6):
                overworkCount = checkoutTime.timestamp() - checkinTime.timestamp()
            else:
                overworkCount = checkoutTime.timestamp() - overworkTime.timestamp()
            if overworkCount < 0 :
                self.worktime = 0
            else : 
                self.worktime = overworkCount
        else :
            self.worktime = 0

        if self.worktime > 0 : 
            self.worktimeMin = round(self.worktime/60,0)

        return self

class PersonRecord :
    name = ""
    depart = ""
    total = 0
    totalCount = 0
    records = []
    

wb = xlrd.open_workbook(filename=source)
detail = wb.sheet_by_index(1)
records = {}
recordsOrg = {}
personRecords = {}

row = 3 
col = 0

# match records
while  row < detail.nrows :
    #col = 0
    #while col < detail.ncols : 
    #    print(detail.cell_value(row,col))
    #    print(" ")
    #    col = col +1
    date = detail.cell_value(row,0).replace("/","-")
    name = detail.cell_value(row,1)
    depart = detail.cell_value(row,3)
    ctype =  "checkout"  if (detail.cell_value(row,5) == "下班打卡") else "checkin"
    ctime = detail.cell_value(row,6)
    ckdate = date + " "+ctime
    if ctime != "" and ctime != "--" :
        t = datetime.fromisoformat( ckdate).isoformat(" ")
    else:
        t = ""

    key = name+"_"+date

    r = records.get(key)
    if r == None :
        r = Record()
        r.name = name 
        r.date = date
        r.depart = depart

    if ctype == "checkin" :
        r.checkin = t
    else :
        r.checkout = t 
    
    records[key] = r
    #print("\n")
    row=row +1



print( "- " + str(row - 3)+" rows processed")


# data filter
recordsCopy = copy.copy(records)
keys = recordsCopy.keys()
for  k in keys: 
    checkinTime = records[k].checkin
    #print(checkinTime)
    if  dateInRange(checkinTime) != True : 
        del records[k]

print("- date filter "+ str(dateStart) + " to " + str(dateEnd))



# stat
for r in records.values() :
    r.updateWorktime()
    key = r.name
    personRecord = personRecords.get(key)
    if personRecord == None : 
        personRecord = PersonRecord()
        personRecord.name = r.name
        personRecord.depart  = r.depart
    
    personRecord.records.append(r)
    personRecord.total = r.worktimeMin + personRecord.total
    personRecord.totalCount = personRecord.totalCount + 1
    personRecords[key] = personRecord


# for r in personRecords.values() :
#     print(r.name + ",t="+str(r.total) + ",tc=" + str(r.totalCount))

## order
keys = sorted(records.keys())



wb = xlwt.Workbook()
ws = wb.add_sheet('打卡明细')

ws.write(0,0,"姓名")
ws.write(0,1,"部门")
ws.write(0,2,"上班时间")
ws.write(0,3,"下班时间")
ws.write(0,4,"加班时长(分)")
row = 1
for k in keys : 
    r = records[k]
    ws.write(row,0,r.name)
    ws.write(row,1,r.depart)
    ws.write(row,2,r.checkin)
    ws.write(row,3,r.checkout)
    ws.write(row,4,r.worktimeMin)
    row = row +1


ws = wb.add_sheet('餐补统计')

ws.write(0,0,"姓名")
ws.write(0,1,"部门")
ws.write(0,2,"加班时长")
ws.write(0,3,"加班天数")
row = 1

sortedItems  = sorted(personRecords.items(),key=lambda x:x[1].total,reverse=True)

for k,r in sortedItems : 
    
    ws.write(row,0,r.name)
    ws.write(row,1,r.depart)
    ws.write(row,2,r.total)
    ws.write(row,3,r.totalCount)
    row = row +1

wb.save(resultFile)




## 输出EXCEL 
# fp = open(resultFile,'w')
# fp.write("<table><tr><th>姓名</th><th>打卡时间</th><th>下班时间</th><th>加班时长</th>\n")
# for k in keys : 
#     r = records[k]
    
#     if r.worktimeMin > 60 : 
#         fp.write("<tr>")
#         fp.write("<td>"+r.name+"</td>")
#         fp.write("<td>"+r.checkin+"</td>")
#         fp.write("<td>"+r.checkout+"</td>")
#         fp.write("<td>"+str(r.worktimeMin)+"</td>")
#         fp.write("</tr>\n")
# fp.write("</table>")
# fp.flush()
# print( "- result saved to " + resultFile)