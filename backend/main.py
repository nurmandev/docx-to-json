from flask import Flask, request, jsonify
from flask_cors import CORS
from pymongo import MongoClient
import boto3
import os
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
import logging
from datetime import datetime
import json

# Load environment variables from .env file
load_dotenv()

# logging.basicConfig(level=logging.DEBUG) 

app = Flask(__name__)
CORS(app)  # To handle cross-origin requests

# MongoDB setup using environment variable for connection string
MONGO_URI = os.getenv('MONGO_URI')
client = MongoClient(MONGO_URI)
db = client['mydatabase']
collection = db['mycollection']

# AWS S3 setup
s3 = boto3.client(
    's3',
    aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY'),
    region_name=os.getenv('AWS_REGION')
)
S3_BUCKET = os.getenv('AWS_S3_BUCKET_NAME')

def save_json_to_file(data):
    try:
        # Define the filename with a timestamp to avoid overwriting
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        file_name = f"stored_data_{timestamp}.json"
        
        # Write JSON data to the file
        with open(file_name, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        
        logging.info(f"Data saved to {file_name}")
    except Exception as e:
        logging.error(f"Error writing JSON data to file: {str(e)}", exc_info=True)

# Route to store JSON data in MongoDB
@app.route('/api/store-json', methods=['POST'])
def store_json():
    try:
        data = request.get_json()  # Get the JSON data from the request
        if not data:
            return jsonify({"error": "No data provided"}), 400

        # Insert the data into MongoDB
        result = collection.insert_one(data)

        # save_json_to_file(data)

        return jsonify({"message": "Data stored successfully", "id": str(result.inserted_id)}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Route to upload image to AWS S3
@app.route('/api/upload-image', methods=['POST'])
def upload_image():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file part"}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No selected file"}), 400
        print(file.filename)
        # Secure the file name
        filename = secure_filename(file.filename)
        print(filename)

        # Upload the image to S3
        s3.upload_fileobj(
            file,
            S3_BUCKET,
            filename,
            # ExtraArgs={"ACL": "public-read"}  # Set the file to be publicly accessible -
        )


        # Generate the file URL
        file_url = f"https://{S3_BUCKET}.s3.{os.getenv('AWS_REGION')}.amazonaws.com/{filename}"

        return jsonify({"message": "Image uploaded successfully", "url": file_url}), 200
    except Exception as e:
        logging.error("Error uploading image to S3", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
