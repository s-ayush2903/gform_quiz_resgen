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

progressive = False
@app.route('/', methods=['GET','POST'])
def file():
   canSendEmail = False
   global correctPoints, incorrectPoints
   if request.method == 'POST':
      if 'files[]' not in request.files:
          flash('No file part')
          return redirect(request.url)
      files = request.files.getlist('files[]')

      print(f"progressive#1: {customUtils.progressive}")
      # Marks field- Number or empty | if number => int (correctly working) | if empty then wo usko as a string read kar raha hai => int mein cast kar rahe hain
      print(type(request.form['pos']))
      # if (isinstance(request.form['pos'], int)) or progressive:
          # correctPoints = int(request.form['pos'])
      # else: 
          # correctPoints = ""
      correctPoints = (int(request.form['pos']) if request.form['pos'] != "" else customUtils.cachedPm)
      incorrectPoints = (int(request.form['neg']) if request.form['neg'] != "" else customUtils.cachedNm)
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
         canSendEmail = True
         customUtils.progressive = False
         print(f"progressive#2: {customUtils.progressive}")
         customUtils.mainFn(correctPoints, incorrectPoints)
         flash('RN wise done')

      if "concise" in request.form:
         canSendEmail = True
         print(f"progressive#3: {customUtils.progressive}")
         customUtils.callConcise(correctPoints, incorrectPoints)
         flash('Concise done')

      if "mail" in request.form:
          if canSendEmail:
             rmMap = customUtils.rollEmailMap
             print("Printing rolMap")

             for roll in rmMap:
                 print(roll, rmMap[roll])

             sendmails(rmMap)
             customUtils.progressive = False
             flash('Mails done')
          else:
               print("-------------")
               print("INVALID ENTRY")
               print("-------------")
               flash("Input ALL THE Fields")


   return redirect('/')

mail = Mail(app) # instantiate the mail class

# configuration of mail
app.config['MAIL_SERVER']='stud.iitp.ac.in'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'csxxx20xx@gmail.com'
app.config['MAIL_PASSWORD'] = 'Whatever123'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)

# message object mapped to a particular URL ‘/’
@app.route("/")
def sendmails(rollMailMap):
    ansDir = os.path.join(os.getcwd(), "ans")
    resultDir = os.path.join(ansDir, "result")
    for key in rollMailMap:
        msg = Message("Quiz Result Out", sender="csxxx20xx@gmail.com", recipients=['stvayush@gmail.com'])
        msg.body = f"Dear Student,\nCSXXX 20XX recent paper marks are attached for reference.\n+{correctPoints} Correct, -{incorrectPoints} for wrong."
        resFileName = os.path.join(resultDir, str(key) + ".xlsx")
        with app.open_resource(resFileName) as fp:
            msg.attach(resFileName, "application/xlsx", fp.read())
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

