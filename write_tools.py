import pandas as pd
from ollama import Client
import json
import os
import re
client = Client(host='http://localhost:11434')





def save_to_new_excel(data_json: str,output_name: str) -> str:
    """
    AIが作成したデータ(JSON形式)を,新しいエクセルファイルとして保存するツール
    """
    try:
        if not isinstance(data_json, str):
            data_json = json.dumps(data_json)

        # 1. 前後のクォートを剥ぎ取る
        output_name = output_name.strip("'").strip('"')
        
        # 2. 【追加】拡張子がなければ強制的に付与する
        if not output_name.lower().endswith('.xlsx'):
            output_name += '.xlsx'

        # 3. JSONの文法ミス（' や None）を物理的に直す
        safe_json = data_json.replace("'", '"')
        safe_json = safe_json.replace("None", "null")
        
        # 4. カンマ連打(,,)を null に置換してパース
        cleaned_json = re.sub(r',(?=\s*,)', ',null', safe_json)
        stripped_json = cleaned_json.strip()
        if not (stripped_json.startswith('[') and stripped_json.endswith(']')):
             # 行ごとにバラバラなら、全体を [] で囲む
             cleaned_json = f"[{cleaned_json.replace('][', '],[')}]"
             cleaned_json = cleaned_json.replace('\n', ',')
        raw_data = json.loads(cleaned_json)

        # 5. データ成形（リストでも辞書でも対応）
        if isinstance(raw_data, dict) and "data" in raw_data:
            list_data = raw_data["data"]
        else:
            list_data = raw_data
        df = pd.DataFrame.from_records(list_data)
        total_count = len(df)
        if len(df.columns) >= 2:
            summary_row = pd.DataFrame([{df.columns[0]: "集計", df.columns[1]: total_count}])
            df = pd.concat([df, summary_row], ignore_index=True)
        df.to_excel(output_name, index=False)
        #ボタン付け
        from openpyxl import load_workbook
        wb=load_workbook(output_name)
        ws=wb.active
        ws.auto_filter.ref = ws.dimensions
        wb.save(output_name)
        return f"指示通りにデータをまとめ、{output_name} として保存しました！"
    except Exception as e:
        return f"保存中にエラーが発生しました: {str(e)}"


def get_columns(file_path: str) -> str:
    """
    エクセルを読み込んで、データの要約や（行列数や統計量を）を返すツール
    """
    try:
        df=pd.read_excel(file_path)
        summary = df.describe(include='all').to_string()
        return f"データの概要はこちらです:\n{summary}"
    except Exception as e:
        return f"エラーが発生しました: {e}"
    
def get_data(file:str) -> str:
   """指定されたテキストファイルの内容をそのまま返すツール。"""
   try:
    with open(file,'r',encoding="utf-8") as f:
        content = f.read()
        return content
   except Exception as e:
        return f"読み込み失敗: {str(e)}"
   
