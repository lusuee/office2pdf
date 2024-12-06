import os

from flask import Flask, send_file, request

app = Flask(__name__)

@app.route("/convert")
def convert():
    return send_file('src/index.html')

def main():
    app.run(port=int(os.environ.get('PORT', 80)))

if __name__ == "__main__":
    main()
