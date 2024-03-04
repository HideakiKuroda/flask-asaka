from flask import Flask, request, render_template, redirect, url_for, flash, session, jsonify
from openpyxl import load_workbook
from datetime import datetime, time
import os
from utilities_1 import edit_excel, intake_from_exl, generate_new_filename,print_excel_file,dayshift_overtime_to_excel,endofshift_overtime_to_excel,onduty_overtime_to_excel,print_totalling_file
import logging
import json
from pathlib import Path

app = Flask(__name__)
app.secret_key = 'secret_key8902083508'

# ロガーの設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# ファイルハンドラーの設定
file_handler = logging.FileHandler('app.log', encoding='utf-8')
file_handler.setLevel(logging.INFO)

# フォーマッターの設定
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# ロガーにハンドラーを追加
logger.addHandler(file_handler)

#最初の画面を開く
@app.route('/', methods=['GET', 'POST'])
def index():
	session.clear()
	message = "「年月日」と「勤務」を入力して作成ボタンを押してください。"
	file_created = False
    # templatesフォルダ内のindex.htmlをレンダリングして返す
	return render_template("index.html", message=message, file_created=file_created)

#テンプレートからnew_filenameを作成して保存する
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
	sheet = workbook.active
	sheet['F4'] = session.get('category', '')
	date_obj = datetime.strptime(session.get('date', ''), '%Y-%m-%d')
	formatted_date = date_obj.strftime('%Y年%m月%d日')
	sheet['B4'] = formatted_date
	sheet['C4'] = session.get('weekday', '')
	sheet['Q4'] = session.get('person', '')
    # 新しいファイル名で保存
	workbook.save(new_file_path)
	new_filename=os.path.basename(new_file_path)
	session['file_created'] = True
	session['message'] = "日報の入力・編集が可能な状態です。"
	session['filename'] = new_filename  # 新しいファイル名もセッションに保存
    
	return redirect(url_for('edit_report', filename=new_filename))


#new_filenameで保存されたファイルの編集画面
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
        with open('./static/captain.txt', 'r', encoding='utf-8') as file:
            people = [line.strip() for line in file if line.strip()]
            # logging.info(people)
        return render_template('edit_report.html', date=date,people=people, category=category, weekday=weekday, 
                            selected_person=selected_person, message=message, file_created=file_created,new_filename=new_filename)
    else:
        # セッションデータがない場合、ファイルからデータを読み込む
        excel_data = intake_from_exl(filename)
		#Excelで格納されている文字列（日付）から日付データとして取得
        date = datetime.strptime(excel_data.get('date', ''), '%Y年%m月%d日')
        category = excel_data.get('category', '')
        weekday = excel_data.get('weekday', '')
        selected_person = excel_data.get('person', '')
        opening = excel_data.get('opening', '')
        closed = excel_data.get('closed', '')
        remarks1 = excel_data.get('remarks1', '')
        remarks2 = excel_data.get('remarks2', '')
        message = "ファイルを読み込みました。編集を続けてください。"
        file_created = True
        new_filename = filename
        work_details_json = json.dumps(excel_data['work_details'], default=custom_time_serializer)
        session['filename'] = new_filename 
        # logger.info('カテゴリ: %s', category)
        # captain.txt から船長リストを作成
        with open('./static/captain.txt', 'r', encoding='utf-8') as file:
            people = [line.strip() for line in file if line.strip()]
        # return redirect(url_for('index'))  # セッションにデータがなければリダイレクト
        return render_template('edit_report.html', date=date, people=people, category=category, weekday=weekday, opening = opening, 
							   closed = closed,remarks1=remarks1,remarks2=remarks2,selected_person=selected_person, message=message, 
							   file_created=file_created,new_filename=new_filename,work_details_json=work_details_json)

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
	'remarks1': request.form['remarks1'],
	'remarks2': request.form['remarks2'],
	'work_details' :[]
	}

	for i in range(1, 8):  # 例えば7隻の船舶データがある場合
		# ドロップダウンから選択されたdetailsの値を取得
		details_value = request.form.get(f'details_{i}')
		# 'その他'が選択された場合、対応する手入力フィールドから値を取得
		if details_value == 'その他':
			details_value = request.form.get(f'other_details_{i}', '')  # 手入力フィールドが空の場合、デフォルト値として空の文字列を設定

		work_data = {
			'shipname': request.form.get(f'shipname_{i}'),
			'berth': request.form.get(f'berth_{i}'),
			'details': details_value,  # 更新されたロジックを使用
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
		form_data['work_details'].append(work_data)  # 正しいインデントレベルに修正
	session['form_data'] = form_data    
	filename=session.get('filename')
	pass_filename = os.path.join('dailyWorkReports',filename)
    # logging.info(pass_filename)
    #時間外の計算
	closed = ''
	opening = ''
	closed = form_data['closed']
	opening = form_data['opening']
	# logger.info('closed: %s', closed)   
	if form_data['category'] in ["日勤", "臨時出勤"]:
		if not closed or not opening:
			flash('開局時間と閉局時間を入力してください')
			return redirect(url_for('edit_report', filename=filename))
		dayshift_overtime_to_excel(pass_filename)
	elif form_data['category'] == "当直": 
		if closed:
			flash('閉局時間は入力しないでください。')
			return redirect(url_for('edit_report', filename=filename))
		if not opening:
			flash('開局時間を入力してください.')
			return redirect(url_for('edit_report', filename=filename))
		onduty_overtime_to_excel(pass_filename)
	elif form_data['category'] == "当直明け":
		if not closed:
			flash('閉局時間を入力してください')
			return redirect(url_for('edit_report', filename=filename))
		endofshift_overtime_to_excel(pass_filename)
	result = edit_excel(form_data)
	flash(result)
	#同時に印刷も行う
	#print_excel_file(pass_filename)
	return redirect(url_for('edit_report', filename=filename))

@app.route('/print_file', methods=['POST'])
def print_file():
	file_register()
	filename=session.get('filename')
	pass_filename = os.path.join('dailyWorkReports',filename)
	print_excel_file(pass_filename)
	flash('印刷が終了しました！')
	return redirect(url_for('edit_report', filename=filename))

@app.route('/print_totalling', methods=['POST'])
def print_totalling():
	file_register()
	filename=session.get('filename')
	pass_filename = os.path.join('dailyWorkReports',filename)
	print_totalling_file(pass_filename)
	flash('印刷が終了しました！')
	return redirect(url_for('edit_report', filename=filename))

#作成されたファイルの一覧を取得する
@app.route('/get_reports')
def get_reports():
    reports_dir = 'dailyWorkReports'
    reports = os.listdir(reports_dir)  # ディレクトリ内のファイルとフォルダのリストを取得
    # 必要に応じて、ファイルのみをリストアップするフィルタリングを行う
    return jsonify(reports)

def custom_time_serializer(obj):
    """カスタムシリアライザ関数。datetime.timeオブジェクトを文字列に変換します。"""
    if isinstance(obj, time):
        return obj.strftime('%H:%M:%S')
    raise TypeError(f"Object of type {obj.__class__.__name__} is not JSON serializable")


logging.basicConfig(filename='app.log', level=logging.INFO, 
                    format='%(asctime)s %(levelname)s:%(message)s')

if __name__ == '__main__':
    app.run(debug=True)