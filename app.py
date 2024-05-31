from flask import Flask, request, send_file
from werkzeug.utils import secure_filename
from auto_alt_text_for_pptx import extractImagesFromPPTX, resetImageAlttexts
import os

app = Flask(__name__)

default_page = '''
    <!doctype html>
    <style>
    body {
        background-image: url('https://ep.jhu.edu/wp-content/uploads/2021/01/zoom-option3.jpg');
    }
    </style>
    <title>Upload a PPTX File</title>
    
    <h1 style="text-align: center;">Upload a PPTX File</h1>
    <form style="text-align: center;" method=post enctype=multipart/form-data>
      <input type=file name=file>
      <select name=option>
        <option value="1">List alt-text for all images in .pptx files</option>
        <option value="2">Extract all images from .pptx files</option>
        <option value="3">Get auto-captions from Vertex-AI for images</option>
        <option value="4">Update alt-text for each image in place</option>
        <option value="5">Delete the image files from each (sub)directory</option>
        <option value="6">Delete the record_captions.csv file from each (sub)directory</option>
        <option value="7">Reset the alt-text for each image to empty string</option>
      </select>
      <input type=submit value=Upload>
    </form>'''


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        option = request.form.get('option')
        if file:
            filename = secure_filename(file.filename)
            file.save(filename)
            
            # Add a switch or if-elif statements here to handle the selected option
            # For example:
            if option == '1':
                 #List alt-text for all images in .pptx file
                 print(f"[Option 1 selected]. Not implemented yet. [List alt-text for all images in .pptx file]")
                 return default_page + '''\n
                    <br />
                    <p> NOT IMPLEMENTED YET</p>
                    <br />
                    '''
            elif option == '2':
                pass
            elif option == '3':
                pass
            elif option == '4':
                pass
            elif option == '5':
                pass
            elif option == '6':
                pass
            elif option == '7':
                resetImageAlttexts(filename)
                return send_file(filename, as_attachment=True)
    return default_page

if __name__ == '__main__':
    app.run(debug=True)