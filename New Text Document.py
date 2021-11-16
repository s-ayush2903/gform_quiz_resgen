from flask import Flask, render_template, request, redirect, flash
from werkzeug.utils import secure_filename
app = Flask(__name__)

@app.route('/')
def upload_form():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        f.save(secure_filename(f.filename))
        #return 'File uploaded successfully. Go back....'
        return redirect('/')

		
if __name__ == '__main__':
   app.run()