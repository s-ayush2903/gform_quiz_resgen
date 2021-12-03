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
   global correctPoints, incorrectPoints, email, password
   if request.method == 'POST':
      finfo = False
      rejForm = 'application/octet-stream' 
      rf = str(request.files)
      if 'files[]' not in request.files:
          flash('No file part')
          return redirect(request.url)
      #if ('files[]' not in rf) or ('octet-stream' in rf) and (not os.listdir(UPLOAD_FOLDER)):
          #flash('Please upload files')
          #return redirect("/")

      files = request.files.getlist('files[]')
      if rejForm not in str(files):
         if len(files) == 2:
            for file in files :
               print(f"{file} || {type(file)}")
               if file and allowed_file(file.filename):
                  filename = secure_filename(file.filename)
                  print(file)
                  print("***********8")
                  file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            flash('File(s) successfully uploaded')
         else: 
            finfo = True
            flash("Please upload the required files!")
            return redirect('/')
      elif not finfo:
         flash("Please upload the required files!")
      #print(f"---files: {files}")
      #print(f"canSendEmails#1: {customUtils.canSendEmails}")

      #print(type(request.form['pos']))


      rp = request.form['pos']
      rn = request.form['neg']

      #print(f"rp: {rp} | rn {rn}")

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

      #print("=============")
      #print(f"{correctPoints}|{incorrectPoints}")
      #print("=============")

      customUtils.cachedPm = correctPoints
      customUtils.cachedNm = incorrectPoints

      #print("-------")
      #print(f"cachedPm: {customUtils.cachedPm}")
      #print("-------")


      if "roll wise" in request.form:
         #print(f"canSendEmails#2: {customUtils.canSendEmails}")
        if len(files) == 2:
            customUtils.canSendEmails = True
            customUtils.mainFn(correctPoints, incorrectPoints)
            flash('Roll Number Wise Marksheet generated')
        else:
            flash('Please upload the required files!')

      if "concise" in request.form:
         #print(f"canSendEmails#3: {customUtils.canSendEmails}")
        if os.path.exists(customUtils.ansDir) and os.path.exists(customUtils.fle) and os.path.exists(customUtils.master):
            customUtils.canSendEmails = True
            customUtils.callConcise(correctPoints, incorrectPoints)
                #print("Printing rolMap")
            flash('Concise Marksheet generated')
        else:
            print("-------------")
            print("INVALID ENTRY")
            print("-------------")
            flash("Please generate Roll Number Wise Marksheet First!")

      if "mail" in request.form:
        if email == "" and password == "":
            flash("Please enter your email and password in code")
            
        else:
            if os.path.exists(customUtils.ansDir) and customUtils.canSendEmails:
                rmMap = customUtils.rollEmailMap
                #print("Printing rolMap")

                sendmails(rmMap)
                flash('Mails done')
                customUtils.canSendEmails = False
            else:
               #print("-------------")
               #print("INVALID ENTRY")
               #print("-------------")
              flash("Please generate Roll Number Wise Marksheet First!")

   return redirect('/')

mail = Mail(app) # instantiate the mail class

# configuration of mail
email = ""
password = "" 
app.config['MAIL_SERVER']='stud.iitp.ac.in'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = email # enter your email here
app.config['MAIL_PASSWORD'] = password # enter your password here
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)

# message object mapped to a particular URL ‘/’
@app.route("/")
def sendmails(rollMailMap):
    ansDir = os.path.join(os.getcwd(), "ans")
    resultDir = os.path.join(ansDir, "result")
    for key in rollMailMap:
        msg = Message("Quiz Result Out", sender=email, recipients=[rollMailMap[key]])
        msg.body = f"Dear Student,\nCSXXX 20XX recent paper marks are attached for reference.\n+{correctPoints} Correct, -{incorrectPoints} for wrong."
        resFileName = os.path.join(resultDir, str(key) + ".xlsx")
        with app.open_resource(resFileName) as fp:
            msg.attach(str(key) + ".xlsx", "application/xlsx", fp.read())
        mail.send(msg)
    return "mails sent"
   # return 'Sent'

if __name__ == "__main__":
    app.run(debug=True)


"""
Blue:  #0000ff
Red:   #ff0000
Green: #008000
Black: #272727
Font:  Century | 12 & 18 font sizes
"""
