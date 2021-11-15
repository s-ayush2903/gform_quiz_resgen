import os      # For File Manipulations like get paths, rename
from flask import Flask, flash, request, redirect, render_template
from werkzeug.utils import secure_filename
from flask_mail import Mail
import shutil
import csv
import openpyxl
from openpyxl.drawing.image import Image

from openpyxl.styles import Alignment, Border, Font, NamedStyle, Side
import os

app=Flask(__name__)
app.secret_key = "secret key" # for encrypting the session
#It will allow below 4MB contents only, you can change it
app.config['MAX_CONTENT_LENGTH'] = 4 * 1024 * 1024
path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')
# Make directory if "uploads" folder not exists
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = set(['csv', 'xlsx'])
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def upload_form():
   return render_template('upload.html')

@app.route('/', methods=['GET','POST'])
def file():
   if request.method == 'POST':
      if 'files[]' not in request.files:
          flash('No file part')
          return redirect(request.url)
      files = request.files.getlist('files[]')

      for file in files:
         if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

      flash('File(s) successfully uploaded')
      global corPoints, incorPoints
      corPoints = request.form['pos']
      incorPoints = request.form['neg']
      pwd = os.getcwd()
      fle = os.path.join(pwd, "uploads\\responses.csv")
      master = os.path.join(pwd, "uploads\\master_roll.csv")
      rootDir = os.path.join(pwd, "ans")
      ansDir = os.path.join(rootDir, "result")

      colls = ["A", "B", "C", "D", "E"]

      ans = []
      absentNameRollMap, concMs, styless = {}, {}, {}


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
          for cul in colls:
              sheet.column_dimensions[cul].width = 18
          sheet.add_image(Image(os.path.join(pwd, "instiLogo.jpeg")), "A1")
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
                          sheet["A10"] = "No."
                          sheet["A11"] = "Marking"
                          sheet["A12"] = "Total"
                          continue
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
                          sheet[col + str(inr)] = int(sheet["B10"].value) * int(sheet["B11"].value)
                      sheet[col + str(inr)].style = getStyle("correct")
                  if col == "C":
                      if inr == 10:
                         sheet[col + str(inr)] = wrong
                      elif inr == 11:
                          sheet[col + str(inr)] = "-" + incorPoints
                      elif inr == 12:
                          sheet[col + str(inr)] = int(sheet["C10"].value) * int(sheet["C11"].value)
                      sheet[col + str(inr)].style = getStyle("incorrect")

          sheet["D10"] = left
          sheet["E10"] = cors + left + wrong
          sheet["E10"].style = getStyle("normal")
          sheet["D11"] = 0
          sheet["D10"].style = getStyle("normal")
          marks = int(cors) * int(corPoints) - int(wrong) * int(incorPoints)
          tmarks = (cors + left + wrong) * int(corPoints)
          mstr=str(marks) + "/" + str(tmarks)
          sheet["E12"] = (mstr
              if not absent
              else "Absent"
          )
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
          for index, conts in enumerate(csv.reader(open(master))):
              if index > 1:
                  if conts[0] not in files:
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


      def mainFn():
          #corPoints = abs(int(input("enter val for correct questions: ")))
          #incorPoints = abs(int(input("enter val for incorrect questions: ")))
          response = prepareResultForPresentStudents()
          if response:
              processLeft()
              prepareConciseMarksheet()
              archiveRes()
              print(os.listdir(ansDir))
          else:
              print("fy and grow up")

      if "roll wise" in request.form:
         prepareResultForPresentStudents()
         flash('RN wise done')
      if "concise" in request.form:
         prepareConciseMarksheet()
         flash('Concise done')
      if "mail" in request.form:
         sendmails()
         flash('Mails done')
      

   return redirect('/')
"""def index():
      if "roll wise" in request.form:
         prepareResultForPresentStudents()
         flash('RN wise done')
      if "concise" in request.form:
         prepareConciseMarksheet()
         flash('Concise done')
      if "mail" in request.form:
         sendmails()
         flash('Mails done')

mail = Mail(app) # instantiate the mail class
   
# configuration of mail
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'yourId@gmail.com'
app.config['MAIL_PASSWORD'] = '*****'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)
   
# message object mapped to a particular URL ‘/’
@app.route("/")
def index():
   msg = Message('Hello',sender ='yourId@gmail.com',recipients = ['receiver’sid@gmail.com'])
   msg.body = 'Hello Flask message sent from Flask-Mail'
   mail.send(msg)
   return 'Sent'
"""
if __name__ == "__main__":
    app.run()
    



"""
Blue:  #0000ff
Red:   #ff0000
Green: #008000
Black: #272727
Font:  Century | 12 & 18 font sizes
"""

