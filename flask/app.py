from flask import Flask,render_template,request
import xlwings as xw
import datetime
import os

app=Flask(__name__)

file_path = os.path.join(os.path.dirname(__file__), "form.xlsx")

if os.path.exists(file_path):
    wb = xw.Book(file_path)  # 既存のファイルを開く
else:
    wb = xw.Book()  # 新しいExcelブックを作成
    wb.save(file_path)  # 新しいファイルを保存
xw.Range("A1").value="時刻"
xw.Range("B1").value="名前"
xw.Range("C1").value="質問"
xw.Range("A2").value=""
xw.Range("B2").value=""
xw.Range("C2").value=""




global count
count=1
@app.route("/",methods=["GET","POST"])
def form():
    global count
    if request.method=="GET":
        return render_template("form.html")
    elif request.method=="POST":
        count+=1
        current_time = datetime.datetime.now().strftime("%H:%M")
        xw.Range((count,1)).value=current_time
        xw.Range((count,2)).value=request.form["name"]
        xw.Range((count,3)).value=request.form["msg"]
        return render_template("on.html")





if __name__=="__main__":
    try:
        # app.run(port=int("5000"),debug=True,host="localhost")
        app.run(host="0.0.0.0", port=5000)
    finally:
        wb.save()
        wb.close()
