from openpyxl import Workbook, load_workbook

wb = load_workbook("PythonExcelPractice.xlsx") #input destination workbook as title
ws = wb.active
name = "" #input name of worker here
info = "" #input information to sort through here

splitString = info.splitlines()
entryNum = 0
entries = [""]*10
for i in range(0, len(splitString)):
    curLine = splitString[i]
    if curLine == "":
        entryNum += 1
    else:
        entries[entryNum] += curLine
        entries[entryNum] += "&&"
entryNum = entryNum + 1
currDate = " "

if name == "Chris C.":
    for i in range(0, entryNum):
        byLine = entries[i].split("&&")
        # last line of each entryis blank!
        # data and address: line one
        dateAddress = byLine[0]
        if dateAddress.find(":") != -1:
            currDate = dateAddress.split(":")[0]
            address = dateAddress.split(":")[1]
        else:
            address = dateAddress
        #time and place
        time = byLine[1]
        hours = ""
        place = time.find("(") + 1
        while time[place] != "h":
            hours += time[place]
            place += 1
        # materials
        if len(byLine) == 5: # has materials included
            materialsLine = byLine[2]
            findSource = materialsLine.split(":")
            matDescript = findSource[1]
            findAmount = matDescript.split("$")
            matDescript = findAmount[0]
            amount = findAmount[1]
            # description
            description = byLine[3]
            ws.append([currDate, address, "Materials", matDescript, name, "", amount])
        else: # no materials
            description = byLine[2]
        ws.append([currDate, address, "Time", description, name, hours])

if name == "Paul C.":
    for i in range(0, entryNum):
        byLine = entries[i].split("&&")
        lines = len(byLine)
        # date
        lineOn = 0
        if byLine[0][0] == 'M' or byLine[0][0] == 'T' or byLine[0][0] == 'W' or byLine[0][0] == 'F' or byLine[0][0] == 'S':
            dateLine = byLine[0]
            place = dateLine.find(" ")
            currDate = dateLine[place:]
            lineOn += 1
        # address
        address = byLine[lineOn]
        lineOn += 1
        # time
        timeLine = byLine[lineOn]
        time = timeLine.split(" ")[1]
        lineOn += 1
        lineOn = len(byLine) - 3
        description = byLine[lineOn]
        lineOn -= 1
        ws.append([currDate, address, "Time", description, name, time])
        while byLine[lineOn].find("$") != -1:
            materialLine = byLine[lineOn]
            materialSplit = materialLine.split("-")
            amount = materialSplit[1]
            if amount.find("$") == -1:
                print("Fix materials here!")
                amount = materialSplit[2]
            materialSource = materialSplit[0]
            ws.append([currDate, address, "Materials", materialSource, name, " ", amount])
            lineOn -= 1

if name == "Ireland C. ":
    for i in range(0, entryNum):
        currEntry = entries[i]
        byLine = currEntry.split("&&")
        currDate = currEntry.split(" ")[0]
        currDate += "/2023"
        currLine = byLine[0]
        place = currLine.find(" ")+1
        address = currLine[place:]
        if currEntry.find("hr") != 1:
            place = currEntry.find("hr") - 2
            hours = ""
            while currEntry[place] != ".":
                hours = currEntry[place] + hours
                place -= 1
            hours = hours[1:]
        else:
            hours = "Not specified"
        while currEntry.find("$") != -1:
            amount = ""
            place = currEntry.find("$") + 1
            while place < len(currEntry)  and currEntry[place] != " ":
                amount += currEntry[place]
                place += 1
            currEntry = currEntry.replace("$", "")
            while amount.find("$") != -1:
                amount = amount.replace("$", "")
            ws.append([currDate, address, "Materials", "Add Source", name, "", amount])
        description = ""
        for j in range(1, len(byLine)):
            description = description + byLine[j]
        end = description.find(hours)
        description = description[:end]
        ws.append([currDate, address, "Time", description, name, hours])

ws.save("PythonExcelPractice.xlsx")

