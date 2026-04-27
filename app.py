from flask import Flask
from flask_cors import CORS


app = Flask(__name__)
CORS(app)


def developing():
    return "功能开发中", 200


@app.route("/api/invoice", methods=["POST"])
def api_invoice():
    return developing()


@app.route("/api/video", methods=["POST"])
def api_video():
    return developing()


@app.route("/api/redemption", methods=["POST"])
def api_redemption():
    return developing()


@app.route("/api/excel", methods=["POST"])
def api_excel():
    return developing()


@app.route("/api/fund", methods=["POST"])
def api_fund():
    return developing()


@app.route("/api/pcf/crawl", methods=["POST"])
def api_pcf_crawl():
    return developing()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
