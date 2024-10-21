from flask import Flask,render_template,request
from werkzeug.utils import secure_filename
from pdfCropMargins import crop
import os
import win32api
import win32print

app = Flask(__name__)

@app.route("/",methods=["GET","POST"])
def hello_world():
    if request.method == "POST":
        try:
            label = request.files["label"]
            if label:
                if 'adjust' in request.form:
                    label.save(os.path.join("C:/Users/Admin/Desktop/AIR Printer/Labels",secure_filename(label.filename)))
                    label= f"C:/Users/Admin/Desktop/AIR Printer/Labels/{secure_filename(label.filename)}"
                    adjusted_label = label.replace(".pdf","_adjusted.pdf")
                    crop(["-p", "0","-u", "-s", label,"-o",adjusted_label])
                    os.startfile(adjusted_label, "print")
                    win32api.ShellExecute(0,"print",adjusted_label,None,".",0)    
                else:    
                    os.startfile(label, "print")
                    win32api.ShellExecute(0,"print",label,None,".",0)    
        except Exception as e:
            print(e)
    return render_template("index.html")

app.run(host="0.0.0.0",port=80,debug=True)
