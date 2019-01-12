from collections import defaultdict
from io import BytesIO

import pandas as pd
from flask import Flask, render_template, request

pd.set_option("display.max_colwidth", -1)


app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


def _to_str(token):
    if type(token) != str:
        token = str(int(token))
    return token


def _hotel_summary(df):
    num_unique_guests = len(df["Token"][df["Token"].notna()].unique())
    end_time = pd.to_datetime(df.groupby(df["Token"])["Date/Time"].max())
    start_time = pd.to_datetime(df.groupby(df["Token"])["Date/Time"].min())
    duration = end_time - start_time

    return {
        "max_duration_token": _to_str(duration.idxmax()),
        "max_duration_minutes": duration.max().seconds // 60,
        "num_unique_guests": num_unique_guests,
    }


def _chat_summary(df):
    df[["Token"]] = df[["Token"]].fillna(method="ffill")

    data = {}
    for token, sub_df in df.groupby("Token"):
        filled = sub_df[["Guest messages", "Bebot messages"]].fillna("")
        filled["Guest messages"] = filled["Guest messages"].str.replace("\n", "<br>")
        filled["Bebot messages"] = filled["Bebot messages"].str.replace("\n", "<br>")
        data[_to_str(token)] = filled.to_html(
            index=False, escape=False, classes="table"
        )

    result = ""
    for t, table in data.items():
        result += f"<h2> {t} </h2>"
        result += table
    return result


@app.route("/result", methods=["POST"])
def generate_result():
    file = request.files["file"]
    if not file.filename.endswith("xlsx"):
        return "Please upload an excel spreadsheet"
    content_bytes = BytesIO(file.read())

    dfs = pd.read_excel(content_bytes, engine="xlrd", sheet_name=None)
    results = {}
    for hotel, df_ in dfs.items():
        try:
            results[hotel] = _hotel_summary(df_)
        except Exception:
            results[hotel] = f"Error: Can't find summary data for {hotel} sheet"
    chats = {}
    for hotel, df_ in dfs.items():
        try:
            chats[hotel] = _chat_summary(df_)
        except Exception:
            chats[hotel] = f"Error: Can't find chat data for {hotel} sheet"

    return render_template("result.html", results=results, chats=chats)


if __name__ == "__main__":
    app.run(debug=True)
