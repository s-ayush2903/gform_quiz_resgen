import customUtils
import os      # For File Manipulations like get paths, rename
from flask import Flask, flash, request, redirect, render_template
from werkzeug.utils import secure_filename
from flask_mail import Mail, Message
import shutil
import csv
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, NamedStyle, Side
import os

path = os.getcwd()
app=Flask(__name__, static_folder=customUtils.baseDir)
app.secret_key = "secret key" # for encrypting the session

#It will allow below 4MB contents only, you can change it
app.config['MAX_CONTENT_LENGTH'] = 4 * 1024 * 1024

# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')

# Make directory if "uploads" folder not exists
if os.path.exists(UPLOAD_FOLDER):
    shutil.rmtree(UPLOAD_FOLDER)
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
   global correctPoints, incorrectPoints
   if request.method == 'POST':
      rf = str(request.files)
      if ('files[]' not in rf) or ('octet-stream' in rf) and (not os.listdir(UPLOAD_FOLDER)):
          flash('No file part')
          return redirect("/")

      files = request.files.getlist('files[]')
      print(f"---files: {files}")
      print(f"canSendEmails#1: {customUtils.canSendEmails}")

      print(type(request.form['pos']))


      rp = request.form['pos']
      rn = request.form['neg']

      print(f"rp: {rp} | rn {rn}")

      if '.' in request.form['pos']:
          correctPoints = float(rp).__round__(2)
      elif rp != "":
          correctPoints = int(rp)
      else:
         correctPoints = customUtils.cachedPm

      if '.' in request.form['neg']:
          incorrectPoints = float(rn).__round__(2)
      elif rn != "":
          incorrectPoints = int(rn)
      else: 
         incorrectPoints = customUtils.cachedNm

      print("=============")
      print(f"{correctPoints}|{incorrectPoints}")
      print("=============")

      customUtils.cachedPm = correctPoints
      customUtils.cachedNm = incorrectPoints

      print("-------")
      print(f"cachedPm: {customUtils.cachedPm}")
      print("-------")

      for file in files:
         if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            flash('File(s) successfully uploaded')

      if "roll wise" in request.form:
         customUtils.canSendEmails = True
         print(f"canSendEmails#2: {customUtils.canSendEmails}")
         customUtils.mainFn(correctPoints, incorrectPoints)
         flash('Roll Number Wise Marksheet generated')

      if "concise" in request.form:
         customUtils.canSendEmails = True
         print(f"canSendEmails#3: {customUtils.canSendEmails}")
         customUtils.callConcise(correctPoints, incorrectPoints)
         flash('Concise Marksheet generated')

      if "mail" in request.form:
          if os.path.exists(customUtils.rootDir):
             rmMap = customUtils.rollEmailMap
             print("Printing rolMap")

             for roll in rmMap:
                 print(roll, rmMap[roll])

             sendmails(rmMap)
             flash('Mails done')
             customUtils.canSendEmails = False
          else:
               print("-------------")
               print("INVALID ENTRY")
               print("-------------")
               flash("Please generate result first!")


   return redirect('/')

mail = Mail(app) # instantiate the mail class

# configuration of mail
senda = "" # enter your email address here
app.config['MAIL_SERVER']='stud.iitp.ac.in'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = senda
app.config['MAIL_PASSWORD'] = '' # enter your password here
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)

# message object mapped to a particular URL ‘/’
@app.route("/")
def sendmails(rollMailMap):
    ansDir = os.path.join(os.getcwd(), "ans")
    resultDir = os.path.join(ansDir, "result")
    for key in rollMailMap:
        msg = Message("Quiz Result Out", sender=senda, recipients=[rollMailMap[key]])
        msg.body = f"Dear Student,\nCSXXX 20XX recent paper marks are attached for reference.\n+{correctPoints} Correct, -{incorrectPoints} for wrong."
        resFileName = os.path.join(resultDir, str(key) + ".xlsx")
        with app.open_resource(resFileName) as fp:
            msg.attach(str(key) + ".xlsx", "application/xlsx", fp.read())
        mail.send(msg)
    return "mails sent"
   # return 'Sent'

if __name__ == "__main__":
    app.run()


"""
Blue:  #0000ff
Red:   #ff0000
Green: #008000
Black: #272727
Font:  Century | 12 & 18 font sizes
"""
