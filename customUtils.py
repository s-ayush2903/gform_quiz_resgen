import csv
import os
import shutil
import openpyxl
from openpyxl.drawing.image import Image

from openpyxl.styles import Alignment, Border, Font, NamedStyle, Side

colls = ["A", "B", "C", "D", "E"]
ans = []

pwd = os.getcwd()

# stores insti logo
baseDir = os.path.join(pwd, "assets")

uploadDir = os.path.join(pwd, "uploads")
fle = os.path.join(uploadDir, "responses.csv")
master = os.path.join(uploadDir, "master_roll.csv")
rootDir = os.path.join(pwd, "ans")
ansDir = os.path.join(rootDir, "result")

absentNameRollMap, concMs, styless = {}, {}, {}
canSendEmails = False
rollWiseDone = False
global cachedNm, cachedPm, rollEmailMap
rollEmailMap = {}

def getStyle(style):
    bd = Side(style="thin")
    if style not in styless.keys():
        baseStyle = NamedStyle(name=f"{style}Style")
        getColor, align = "ff000000", "center"
        isTitle = False
        if style is "correct":
            getColor = "008000"
        elif style is "incorrect":
            getColor = "ff0000"
        elif style is "absolute":
            getColor = "0000ff"
        elif style is "normal":
            align = "center"
        elif "title" in style:
            if style[0] is "l":
                isTitle = True
                align = "left"
            elif style[0] is "r":
                align = "right"
            elif style[0] is "m":
                isTitle = True

        baseStyle.font = Font(name="Century", size=12, bold=isTitle, color=getColor)
        baseStyle.alignment = Alignment(horizontal=align)
        if style != "ltitle" and style != "rtitle":
            baseStyle.border = Border(bd, bd, bd, bd)
        styless[style] = baseStyle
    # print(f"styless: {styless}")
    return styless[style]


def prepareQuizResult(rollNo, line=[], absent=False):
    global cors, left, wrong
    cors, left, wrong = 0, 0, 0
    wb = openpyxl.Workbook()
    sheet = wb.active
    fileName = os.path.join(ansDir, rollNo + ".xlsx")

    if not absent:
        rollEmailMap[rollNo] = [line[1], line[4]]

    for cul in colls:
        sheet.column_dimensions[cul].width = 18
    sheet.add_image(Image(os.path.join(baseDir, "instiLogo.jpeg")), "A1")
    sheet.merge_cells("A5:E5")
    sheet["A5"] = "Marksheet"
    sheet["A5"].font = Font(name="Century", size=18, bold=True, underline="single")
    sheet["A5"].alignment = Alignment(horizontal="center")
    sheet["A6"] = "Name:"
    sheet["A7"] = "Roll Number:"
    sheet["A6"].style = getStyle("rtitle")
    sheet["A7"].style = getStyle("rtitle")
    sheet["B6"] = line[3] if not absent else absentNameRollMap[rollNo]
    sheet["B7"] = line[6] if not absent else rollNo
    sheet["B6"].style = getStyle("ltitle")
    sheet["B7"].style = getStyle("ltitle")
    sheet["D6"] = "Exam:"
    sheet["D6"].style = getStyle("rtitle")
    sheet["E6"] = "quiz"
    sheet["E6"].style = getStyle("ltitle")
    sheet.append([""])
    sheet.append(["", "Right", "Wrong", "Not Attempt", "Max"])
    for cell in sheet["9:9"]:
        cell.style = getStyle("mtitle")
    qCount, rowNum = 30, 15

    colL, colR = "A", "B"
    onceCompleted = False

    lst = line[7:] if not absent else ans
    for ind, val in enumerate(lst):
        temp = val.strip() if not absent else ""
        if rowNum > 40 or rowNum is 15:
            if onceCompleted:
                colR = chr(ord(colR) + 3)
                colL = chr(ord(colR) - 1)
            sheet[colL + "15"] = "Student Ans"
            sheet[colR + "15"] = "Correct Ans"
            sheet[colL + "15"].style = getStyle("mtitle")
            sheet[colR + "15"].style = getStyle("mtitle")
            onceCompleted = True
            rowNum = str(16)
        sheet[str(colR + str(rowNum))] = str(ans[ind])
        sheet[colL + str(rowNum)] = temp
        sheet[colR + str(rowNum)].style = getStyle("absolute")

        if temp == ans[ind]:
            sheet[colL + str(rowNum)].style = getStyle("correct")
            cors += 1
        elif not temp:
            sheet[colL + str(rowNum)].style = getStyle("normal")
            left += 1
        else:
            sheet[colL + str(rowNum)].style = getStyle("incorrect")
            wrong += 1
        rowNum = int(rowNum) + 1

    for inr in range(10, 13):
        for col in colls:
            if col == "A" or col == "D" or col == "E":
                if col == "A":
                    if inr == 10:
                        sheet[col + str(inr)] = "No."
                    if inr == 11:
                        sheet[col + str(inr)] = "Marking"
                    if inr == 12:
                        sheet[col + str(inr)] = "Total"
                if col == "D" and inr != 9:
                    sheet[col + str(inr)].style = getStyle("normal")
                else:
                    sheet[col + str(inr)].style = getStyle("mtitle")
            if col == "B":
                if inr == 10:
                    sheet[col + str(inr)] = cors
                elif inr == 11:
                    sheet[col + str(inr)] = corPoints
                elif inr == 12:
                    sheet[col + str(inr)] = (sheet["B10"].value) * (
                        sheet["B11"].value
                    )
                sheet[col + str(inr)].style = getStyle("correct")
            if col == "C":
                if inr == 10:
                    sheet[col + str(inr)] = wrong
                elif inr == 11:
                    sheet[col + str(inr)] = -1 * incorPoints
                    # sheet[col + str(inr)] = "-" + str(incorPoints)
                elif inr == 12:
                    sheet[col + str(inr)] = (sheet["C10"].value) * (sheet["C11"].value)
                sheet[col + str(inr)].style = getStyle("incorrect")

    sheet["D10"] = left
    sheet["E10"] = cors + left + wrong
    sheet["E10"].style = getStyle("normal")
    sheet["D11"] = 0
    sheet["D10"].style = getStyle("normal")
    marks = (cors) * (corPoints) - (wrong) * (incorPoints)
    tmarks = (cors + left + wrong) * (corPoints)
    mstr = str(marks) + "/" + str(tmarks)
    sheet["E12"] = mstr if not absent else "Absent"
    sheet["E12"].style = getStyle("absolute")
    concMs[rollNo] = str(
        str(cors * corPoints)
        + ","
        + str(wrong * incorPoints)
        + ","
        + sheet["E12"].value
    )
    sheet.title = "quiz"
    wb.save(fileName)


def prepareResultForPresentStudents() -> bool:
    if os.path.exists(ansDir):
        shutil.rmtree(ansDir)
    os.makedirs(ansDir)

    for index, line in enumerate(csv.reader(open(fle))):
        if index == 1:
            if line[6] == "ANSWER":
                for _ in line[7:]:
                    ans.append(_.strip())
            else:
                print("fy")
                return False
        fileName = os.path.join(ansDir, line[6] + ".xlsx")
        if index > 1:
            prepareQuizResult(line[6], line)
    return True

def processLeft():
    files = os.listdir(ansDir)
    print("_______________________")
    print("_______________________")
    for index, conts in enumerate(csv.reader(open(master))):
        if index > 1:
            if f"{conts[0].upper()}.xlsx" not in files:
                print(f"noo, not found: {conts[0].upper()}")
                absentNameRollMap[conts[0]] = conts[1]
                prepareQuizResult(conts[0], absent=True)


def prepareConciseMarksheet():
    concMsFile = os.path.join(rootDir, "concise_marksheet.csv")
    if os.path.exists(concMsFile):
        os.remove(concMsFile)
    with open(concMsFile, "w") as cmfObj:
        cmfObj.write("Roll,positive_marks,negative_marks,total_marks")
        for roll in concMs:
            cmfObj.write("\n")
            cmfObj.write(str(roll + "," + concMs[roll]))
    return concMs


def archiveRes():
    shutil.make_archive("result", "zip", rootDir)
    return True


def mainFn(cpts, incPts):
    global corPoints, incorPoints
    corPoints = cpts
    incorPoints = incPts
    response = prepareResultForPresentStudents()
    if response:
        rollWiseDone = True
        # processLeft()
        # prepareConciseMarksheet()
        # archiveRes()
        # print(os.listdir(ansDir))
    else:
        return false
        print("fy and grow up")

def callConcise(cpts, incPts):
    if not rollWiseDone:
        mainFn(cpts, incPts)
    processLeft()
    prepareConciseMarksheet()
    archiveRes()
    return True

