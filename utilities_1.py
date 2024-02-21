from flask import Flask, request, request,flash, redirect, url_for
from openpyxl import load_workbook
from flask import session
import os

#app.py で form_dataに格納されたデータをdata としてExcelファイルに書き込む
def edit_excel(data):
    exfilename = session.get('filename')
    if not exfilename:
        return 'Filename is not set in session.'
    exfilename = os.path.join('dailyWorkReports', exfilename)
    book = load_workbook(exfilename)
    sheet = book.active
    # セルにデータを挿入
    sheet['F4'] = data['category']
    sheet['B4'] = data['date']
    sheet['C4'] = data['weekday']
    sheet['Q4'] = data['person']
    sheet['C22'] = data['opening']
    sheet['C23'] = data['closed']

    # 船舶詳細データの書き込み
    row = 8  # 開始行
    for work_detail in data['work_details']:
        if row > 20:
            break  # 20行を超えたらループを終了
        sheet[f'A{row}'] = work_detail.get('shipname')
        sheet[f'E{row}'] = work_detail.get('berth')
        sheet[f'F{row}'] = work_detail.get('details')
        sheet[f'G{row}'] = work_detail.get('schedule')
        sheet[f'H{row}'] = work_detail.get('departure')
        sheet[f'I{row}'] = work_detail.get('onsite')
        sheet[f'J{row}'] = work_detail.get('start')
        sheet[f'K{row}'] = work_detail.get('end')
        sheet[f'L{row}'] = work_detail.get('arrival')
        sheet[f'M{row}'] = work_detail.get('usage')
        sheet[f'N{row}'] = work_detail.get('certificate')
        sheet[f'N{row-1}'] = work_detail.get('partner')
        row += 2  # 次の入力行を設定（一行飛ばし）

    # パートナー情報の書き込み
    # row = 7  # 開始行（パートナー情報はN列に入力）
    # for work_detail in data['work_details']:
    #     if row > 19:
    #         break
    #     sheet[f'N{row}'] = work_detail.get('partner')
    #     row += 2  # 次の入力行を設定（一行飛ばし）

    # 変更を保存
    book.save(exfilename)

    return f'Excel ファイル {exfilename} に書き込みが完了しました.'

#Excelファイルに書き込まれたデータを読み込む
def intake_from_exl(filename):
    exfilename = filename
    if not exfilename:
        return 'Filename is not set in session.'
    exfilename = os.path.join('dailyWorkReports', exfilename)
    book = load_workbook(exfilename)
    sheet = book.active

    excel_data = {
    'date' : sheet['B4'].value,
    'weekday' : sheet['C4'].value,
    'category' : sheet['E4'].value,
    'person' : sheet['Q4'].value,
    'opening': sheet['C22'].value,
    'closed' : sheet['C23'].value,
    'work_details' :[]
    }
    row = 8
    while row <= 20:
        shipname = sheet[f'A{row}'].value
        if not shipname:  # 船名がなければループを終了
            break
        work_detail = {
            'shipname': shipname,
            'berth': sheet[f'E{row}'].value,
            'details': sheet[f'F{row}'].value,
            'schedule': sheet[f'G{row}'].value,
            'departure': sheet[f'H{row}'].value,
            'onsite': sheet[f'I{row}'].value,
            'start': sheet[f'J{row}'].value,
            'end': sheet[f'K{row}'].value,
            'arrival': sheet[f'L{row}'].value,
            'usage': sheet[f'M{row}'].value,
            'certificate': sheet[f'N{row}'].value,
            'partner': sheet[f'N{row-1}'].value  # パートナー情報は1行下
        }
        excel_data['work_details'].append(work_detail)
        row += 2  # 次の船舶情報へ（一行飛ばし）
        book.save(exfilename)
    return excel_data

def generate_new_filename(base_path):
    # ファイルの基本名と拡張子を分離
    base, extension = os.path.splitext(base_path)
    counter = 1  # 連番の開始

    # 新しいファイル名を生成
    new_file_path = f"{base}({counter}){extension}"

    # 生成したファイル名が既に存在する場合は、連番を増やして再試行
    while os.path.exists(new_file_path):
        flash('ファイルが既に存在します！') 
        counter += 1
        new_file_path = f"{base}({counter}){extension}"

    return new_file_path

def save_template_as_new_file(date, template_path):
    # テンプレートファイルを読み込む
    book = load_workbook(template_path)
    sheet = book.active
    
    # 特定のセルから情報を取得する（例: O3セルから作業場所を取得）
    work_location = sheet['O3'].value  # 作業場所等が記されていると仮定
    
    # 新しいファイル名を生成（例: "2023-01-01_東京工場_作業日報.xlsx"）
    new_file_name = f"{date}_{work_location}_作業日報.xlsx"
    
    # 新しいファイル名でテンプレートを保存
    book.save(new_file_name)
    return new_file_name  # 生成された新しいファイル名を返す
