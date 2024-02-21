from flask import Flask, request, render_template, request, redirect, url_for, flash, session, jsonify
from openpyxl import load_workbook
import os
from utilities_1 import edit_excel, intake_from_exl, generate_new_filename
import logging
import json

app = Flask(__name__)
app.secret_key = 'secret_key8902083508'

@app.route('/', methods=['GET', 'POST'])
def index():
    message = "「年月日」と「勤務」を入力して作成ボタンを押してください。"
    file_created = False
    # templatesフォルダ内のindex.htmlをレンダリングして返す
    return render_template("index.html", message=message, file_created=file_created)

@app.route('/create_report', methods=['GET', 'POST'])
def create_report():
	session['date'] = request.form['date']
	session['category'] = request.form['category']
	session['weekday'] = request.form['weekday']
	session['person'] = request.form['person']
    
    # 日付またはカテゴリのいずれかが入力されていない場合
	if not session['date'] or not session['category']:
		flash('入力がありません。日付とカテゴリを両方入力してください。') 
		return redirect(url_for('index'))
    
	new_filename = f"{session['date']}「あさか丸」{session['category']}ー作業日報.xlsx"
    # テンプレートファイルのパス
	template_path = 'adm_template.xlsx'
    # 新しいファイルの保存先パス
	new_file_path = generate_new_filename(os.path.join('dailyWorkReports', new_filename))
    # テンプレートファイルを読み込む
	workbook = load_workbook(template_path)
    # 新しいファイル名で保存
	workbook.save(new_file_path)
	new_filename=os.path.basename(new_file_path)
	session['file_created'] = True
	session['message'] = "日報の入力・編集が可能な状態です。"
	session['filename'] = new_filename  # 新しいファイル名もセッションに保存
    
	return redirect(url_for('edit_report', filename=new_filename))

@app.route('/edit/<filename>', methods=['GET', 'POST'])
def edit_report(filename):
    if 'file_created' in session:
    # セッションからデータを取り出す
        date = session.get('date', '')
        category = session.get('category', '')
        weekday = session.get('weekday', '')
        selected_person = session.get('person', '')
        message = session.get('message', '')
        file_created = session.get('file_created', False)
        new_filename = session.get('filename')
        session.pop('file_created', None) 
        # return redirect(url_for('index'))  # セッションにデータがなければリダイレクト
        return render_template('edit_report.html', date=date,people=people, category=category, weekday=weekday, 
                            selected_person=selected_person, message=message, file_created=file_created,new_filename=new_filename)
    else:
        # セッションデータがない場合、ファイルからデータを読み込む
        excel_data = intake_from_exl(filename)
        date = excel_data.get('date', '')
        category = excel_data.get('category', '')
        weekday = excel_data.get('weekday', '')
        selected_person = excel_data.get('person', '')
        opening = excel_data.get('opening', '')
        closed = excel_data.get('closed', '')
        message = "ファイルを読み込みました。編集を続けてください。"
        file_created = True
        new_filename = filename
        work_details_json = json.dumps(excel_data['work_details'])
        session['filename'] = new_filename 
        # logging.info(work_details_json)
        # captain.txt から船長リストを作成
        # return redirect(url_for('index'))  # セッションにデータがなければリダイレクト
        return render_template('edit_report.html', date=date,people=people, category=category, weekday=weekday, opening = opening, closed = closed,
                                selected_person=selected_person, message=message, file_created=file_created,new_filename=new_filename,work_details_json=work_details_json)
with open('./static/captain.txt', 'r', encoding='utf-8') as file:
    people = [line.strip() for line in file if line.strip()]

# 入力データをexcelファイルに書き込みするためform_dataに格納
@app.route('/register', methods=['POST'])
def file_register():
    form_data = {
    'date' : request.form['date'],
    'weekday' : request.form['weekday'],
    'category' : request.form['category'],
    'person' : request.form['person'],
    'closed' : request.form['closed'],
    'opening': request.form['opening'],
    'work_details' :[]
    }

    for i in range(1, 8):  # 例えば7隻の船舶データがある場合
        work_data = {
            'shipname': request.form.get(f'shipname_{i}'),
            'berth': request.form.get(f'berth_{i}'),
            'details': request.form.get(f'details_{i}'),
            'schedule': request.form.get(f'schedule_{i}'),
            'departure': request.form.get(f'departure_{i}'),
            'onsite': request.form.get(f'onsite_{i}'),
            'start': request.form.get(f'start_{i}'),
            'end': request.form.get(f'end_{i}'),
            'arrival': request.form.get(f'arrival_{i}'),
            'usage': request.form.get(f'usage_{i}'),
            'partner': request.form.get(f'partner_{i}'),
            'certificate': request.form.get(f'certificate_{i}'),
        }
        form_data['work_details'].append(work_data)
    session['form_data'] = form_data    
    result = edit_excel(form_data)
    flash(result)
    return redirect(url_for('edit_report', filename=session.get('filename')))

@app.route('/get_reports')
def get_reports():
    reports_dir = 'dailyWorkReports'
    reports = os.listdir(reports_dir)  # ディレクトリ内のファイルとフォルダのリストを取得
    # 必要に応じて、ファイルのみをリストアップするフィルタリングを行う
    return jsonify(reports)

# logging.basicConfig(filename='app.log', level=logging.INFO, 
#                     format='%(asctime)s %(levelname)s:%(message)s')

if __name__ == '__main__':
    app.run(debug=True)