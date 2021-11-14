from flask import Flask, jsonify, request, session, send_file
from werkzeug.utils import secure_filename
import os
import win32com.client

UPLOAD_FOLDER = './files'
IMAGE_FOLDER = './static'
ALLOWED_EXTENSIONS = {'pptx'}

app = Flask(__name__, static_url_path='/images')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['IMAGE_FOLDER'] = IMAGE_FOLDER

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('No file part')
        file = request.files['file']
        if file.filename == '':
            print('No selected file')
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            image_url_array = presentationToImages(os.path.abspath(file_path), 'image')
            os.remove(file_path)
            return jsonify(image_url_array)
    return ''

def presentationToImages(pres_path, new_file_name):
    # print('-----'+file_path)
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Presentation = Application.Presentations.Open(pres_path)
    filename = new_file_name
    rel_img_path = os.path.join(app.config['IMAGE_FOLDER'], filename)
    abs_img_path = os.path.abspath(rel_img_path)

    #saves and converts each slide in pres
    img_url_array = []
    index = 0
    for slide in Presentation.Slides:
        slide_name = filename+str(index)+'.jpg'
        slide_abs_path = abs_img_path+str(index)+'.jpg'
        slide.Export(slide_abs_path, "JPG")
        img_url_array.append(slide_name)
        index += 1

    Application.Quit()
    Presentation =  None
    Application = None
    return img_url_array

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

if __name__ == "__main__":
    app.run(debug=True)