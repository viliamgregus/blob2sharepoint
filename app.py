# Import used Libraries
from flask import Flask
from flask import jsonify

app = Flask(__name__)

@app.route("/")

def run():
    import modules.project as project

    project.from_azure_to_sharepoint()
    resp = jsonify(success=True)
    return resp