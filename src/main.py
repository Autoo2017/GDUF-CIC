import os
import time

import readExcel
import uuid

from flask import Flask, request, render_template, send_from_directory, url_for

app = Flask(__name__)


def generate_uuid():
    userid = str(uuid.uuid4())
    return userid

UUID = generate_uuid()

def get_uuid():
    uuid = UUID
    return uuid


@app.route('/')
def index():
    return render_template('index.html', uuid=get_uuid())


@app.route('/user/<uuid>.xls.ics', methods=['GET'])
def download(uuid):
    filename = uuid + '.xls.ics'
    return send_from_directory('user', filename, as_attachment=True)


@app.route('/user', methods=['POST'])
def upload_file():
    file = request.files['file']
    if file:
        filenames = get_uuid() + '.xls'
        upload_files = os.path.join('./user', filenames)
        file.save(upload_files)

        try:
            readExcel.read(filenames)
        except ImportError as e:#上传了个非法的xls
            return "上传失败（103）"
        except ValueError as e:#上传了不明文件
            return "上传失败（102）"
        except Exception as e:
            return str(e)

        download_url = url_for('download', uuid = get_uuid(),_external=True)
        #delay 6 seconds and then return render
        time.sleep(6)
        return render_template('download.html', url=download_url, uuid=get_uuid())
    else:
        return "上传失败（104）"


if __name__ == '__main__':
    app.run(host = '0.0.0.0', debug=True)
