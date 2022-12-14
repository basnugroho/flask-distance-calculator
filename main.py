from flask import Flask, render_template, request, session, send_file
import pandas as pd
import numpy as np
import os
from io import BytesIO
from werkzeug.utils import secure_filename
import googlemaps
from datetime import datetime
import time
import csv
from functools import reduce
import haversine as hs
from haversine import Unit
 
#*** Flask configuration
 
# Define folder to save uploaded files to process further
UPLOAD_FOLDER = os.path.join('staticFiles', 'uploads')
 
# Define allowed files (for this example I want only csv file)
ALLOWED_EXTENSIONS = {'xlsx'}
 
app = Flask(__name__, template_folder='templates', static_folder='staticFiles')
# Configure upload file path flask
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
 
# Define secret key to enable session
app.secret_key = 'This is your secret key to utilize session in Flask'

quota = 2000

cwd = os.getcwd()
API_key = 'AIzaSyBvJaPeqmGdxeJlxfDwDyNmjj1h1ZD9_bg'   #enter the key you got from Google. I removed mine here
gmaps = googlemaps.Client(key=API_key)

def read_upload_quota():
    df = pd.read_csv(cwd+"/log/processed.csv")
    return df['processed'].sum()

def update_upload_quota(waktu, row, elapsed_time):
    fileName = cwd+"/log/processed.csv"
    fileEmpty = os.stat(fileName).st_size == 0
    with open(fileName, "a") as csvfile:
        headers = ['date_time', 'processed', 'elapsed_time']
        writer = csv.DictWriter(csvfile, delimiter=',', lineterminator='\n',fieldnames=headers)
        if fileEmpty:
            writer.writeheader()  # file doesn't exist yet, write a header

        writer.writerow({'date_time': waktu, 'processed': row, 'elapsed_time': str(elapsed_time)})

def google_calc_walking_distance(origin_point, dest_point):
  try:
    result = gmaps.distance_matrix(origin_point, dest_point, mode='walking')['rows'][0]['elements'][0]['distance']['value']
  except KeyError:
    result = "error"
  return result

def google_calc_distance_bulk(lata,latb,latc,latd):
    distances = []
    for i,x in enumerate(lata):
        origins = (lata[i],latb[i])
        destination = (latc[i],latd[i])
        distances.append(google_calc_walking_distance(origins,destination))
    return distances

def haversine_calc_dists(lata,latb,latc,latd):
    distances = []
    for i,x in enumerate(lata):
      try:
        distances.append(hs.haversine((lata[i],latb[i]),(latc[i],latd[i]),unit=Unit.METERS))
      except:
        distances.append("error")
    return distances

@app.route('/')
def index():
    processed = read_upload_quota()
    return render_template('index_upload_and_show_data.html', processed=processed, quota=quota)
 
@app.route('/',  methods=("POST", "GET"))
def uploadFile():
    if request.method == 'POST':
        # upload file flask
        uploaded_df = request.files['uploaded-file']
 
        # Extracting uploaded data file name
        data_filename = secure_filename(uploaded_df.filename)
 
        # flask upload file to database (defined uploaded folder in static path)
        #uploaded_df.save(os.path.join(app.config['UPLOAD_FOLDER'], data_filename))

        # Storing uploaded file path in flask session
        #session['uploaded_data_file_path'] = "/root/flask-distance-calculator/"+data_filename

        # process the file
        df = pd.read_excel(uploaded_df)
        hav_distances = haversine_calc_dists(df["LATITUDE CUSTOMER"], df["LONGITUDE CUSTOMER"], df["LATITUDE ODP"], df["LONGITUDE ODP"])
        df["haversine_distance (m)"] = hav_distances
        print("calculating, estimated time: "+str(0.3*len(df.index)))
        now = datetime.now()
        date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
        start = time.time()
        distances = google_calc_distance_bulk(df["LATITUDE CUSTOMER"], df["LONGITUDE CUSTOMER"], df["LATITUDE ODP"], df["LONGITUDE ODP"])
        df["walking_distance (m)"] = distances
        end = time.time()
        print(f"elapsed time: {end - start}(s)")
 
        #create an output stream
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')

        df.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Result")
        workbook = writer.book
        # worksheet = writer.sheets["Result"]

        format = workbook.add_format()
        format.set_bg_color('#eeeeee')
        # worksheet.set_column(0,9,28)
        
        #the writer has done its job
        writer.close()

        #go back to the beginning of the stream
        output.seek(0)

        #finally return the file
        update_upload_quota(date_time, len(df["walking_distance (m)"]), "{0:,.2f}".format(end - start))
        print("quota updated")
        df.to_excel(f"{cwd}/results/result_{data_filename}")
        print("result saved")
        return send_file(output, download_name=f"result_{data_filename}.xlsx", as_attachment=True);return render_template('index_upload_and_show_data_page2.html', time_elapsed="{0:,.2f}".format(end - start))
 
@app.route('/show_data')
def showData():
    # Retrieving uploaded file path from session
    data_file_path = session.get('uploaded_data_file_path', None)
 
    # read csv file in python flask (reading uploaded csv file from uploaded server location)
    uploaded_df = pd.read_excel(data_file_path)
 
    # pandas dataframe to html table flask
    uploaded_df_html = uploaded_df.to_html()
    return render_template('timer.html', data_var = uploaded_df_html)

@app.route('/download')
def download():
    #create a random Pandas dataframe
    df_1 = pd.DataFrame(np.random.randint(0,10,size=(10, 4)), columns=list('ABCD'))

    #create an output stream
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    #taken from the original question
    df_1.to_excel(writer, startrow = 0, merge_cells = False, sheet_name = "Sheet_1")
    workbook = writer.book
    worksheet = writer.sheets["Sheet_1"]
    format = workbook.add_format()
    format.set_bg_color('#eeeeee')
    worksheet.set_column(0,9,28)

    #the writer has done its job
    writer.close()

    #go back to the beginning of the stream
    output.seek(0)

    #finally return the file
    return send_file(output, download_name="testing.xlsx", as_attachment=True)
 
if __name__=='__main__':
    app.run(host='0.0.0.0')
