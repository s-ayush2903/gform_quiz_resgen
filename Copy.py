import customUtils
import os      # For File Manipulations like get paths, rename
from flask import Flask, flash, request, redirect, render_template
from werkzeug.utils import secure_filename
# from flask_mail import Mail
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
if  os.path.exists(UPLOAD_FOLDER):
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
   global correctPoints, incorrectPoints
   if request.method == 'POST':
      if 'files[]' not in request.files:
          flash('No file part')
          return redirect(request.url)
      files = request.files.getlist('files[]')

      print(f"progressive#1: {customUtils.progressive}")
      correctPoints = (int(request.form['pos']) if not customUtils.progressive else customUtils.cachedPm)
      incorrectPoints = (int(request.form['neg']) if not customUtils.progressive else customUtils.cachedNm)
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
         customUtils.progressive = True
         print(f"progressive#2: {customUtils.progressive}")
         customUtils.mainFn(correctPoints, incorrectPoints)
         flash('RN wise done')
      if "concise" in request.form:
         customUtils.progressive = False
         print(f"progressive#3: {customUtils.progressive}")
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

