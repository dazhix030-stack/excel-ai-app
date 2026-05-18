from ollama import Client
from mcp.server.fastmcp import FastMCP
from write_tools import get_columns
from write_tools import save_to_new_excel
from write_tools import get_data
import streamlit as st
client = Client(host='http://localhost:11434')

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_data",
            "description": "指定されたテキストファイルの内容を返す",
            "parameters": {
                "type": "object",
                "properties": {"file": {"type": "string", "description": "ファイルパス"}},
                "required": ["file"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "save_to_new_excel",
            "description": "AIが作成したデータをエクセルとして保存する",
            "parameters": {
                "type": "object",
                "properties": {
                    "data_json": {"type": "string"},
                    "output_name": {"type": "string"}
                },
                "required": ["data_json", "output_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_columns",
            "description": "エクセルの要約・統計量を返す",
            "parameters": {
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
                "required": ["file_path"]
            }
        }
    }
]
TOOL_MAP = {
    "get_data": get_data,
    "save_to_new_excel": save_to_new_excel,
    "get_columns": get_columns,
}
mcp=FastMCP("MyExcelToolr")

SYSTEM_PROMPT = """あなたはエクセルのプロです。
指示に応じて、以下のルールでツールを使い分けてください。

1. 【エクセルファイル (.xlsx) の場合】:
   - まず `get_columns` で行列数や統計量を確認し、データの構造を把握してください。
   - 結果を解説してください
   get_columnsをした場合はほかの関数は使用禁止です

2. 【テキストファイル (.txt / .csv) の場合】:
   - `get_columns` は使用禁止です。直接 `get_data` を使って内容を読み取ってください。

3. 【共通ルール】:
   - 計算（合計など）が必要な場合は、`get_data` で取得した数値をもとに AI が計算してください。
   - `save_to_new_excel` を実行する際は、必ずダブルクォートを使用した正しいJSON形式を守ってください。"""
st.title("エクセル集計")
uploaded_files = st.file_uploader("ファイルをアップロードしてください", accept_multiple_files=True)
selected_file = None
if uploaded_files:
    file_names = [f.name for f in uploaded_files]
    selected_file = st.selectbox("処理したいファイルを選んでください", file_names)
    st.write(f"現在 **{selected_file}** が選択されています。")
 # セッション初期化
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "system", "content": SYSTEM_PROMPT}]
if "chat_history" not in st.session_state:
     st.session_state.chat_history = []
for msg in st.session_state.chat_history:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])
#チャットエリア
if query := st.chat_input("AIへの指示を書いて下さい"):
    if selected_file and selected_file.endswith('.xlsx'):
        input_text = f"ファイル「{selected_file}」はエクセルです。まず `get_columns` で概要を確認してから、指示に応えてください：{query}"
    elif selected_file:
        input_text = f"ファイル「{selected_file}」を `get_data` で読み取ってから、指示に応えてください：{query}"
    else:
        input_text = query
    st.session_state.chat_history.append({"role": "user", "content": query})
    st.session_state.messages.append({"role": "user", "content": input_text})
    with st.chat_message("user"):
        st.write(query)

    # AIの処理
    with st.chat_message("assistant"):
        with st.spinner("処理中..."):
            max_iterations = 10
            count = 0
            while count < max_iterations:
                count += 1
                if count == max_iterations:
                    st.error("ループ回数が上限に達しました。処理を中断します。")
                    break
                response = client.chat(
                    model="qwen2.5:7b",
                    messages=st.session_state.messages,
                    tools=TOOLS
                )
                msg = response["message"]

                # ツール呼び出しなし → 最終回答
                if not msg.get("tool_calls"):
                    st.write(msg["content"])
                    st.session_state.chat_history.append({"role": "assistant", "content": msg["content"]})
                    st.session_state.messages.append({"role": "assistant", "content": msg["content"]})
                    break

                # ツール呼び出しあり
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": msg.get("content", ""),
                    "tool_calls": msg["tool_calls"]
                })

                for tool_call in msg["tool_calls"]:
                    name = tool_call["function"]["name"]
                    args = tool_call["function"]["arguments"]

                    st.info(f"{name} を実行中...")
                    if name == "get_data":
                        target_name = args.get("file") or args.get("file_path")
                        file_obj = next((f for f in uploaded_files if f.name == target_name), None) if uploaded_files else None
                        
                        if file_obj:
                            # エクセルファイルだった場合、強引にデコードせず pandas で読み取る
                            if file_obj.name.endswith('.xlsx'):
                                import pandas as pd
                                try:
                                    # エンジンを指定して読み込み、AIが読めるCSV形式に変換
                                    df = pd.read_excel(file_obj, engine='openpyxl')
                                    result = f"このファイルはエクセル形式です。中身を抽出しました：\n{df.to_csv(index=False)}"
                                except Exception as e:
                                    result = f"エクセル読み込み中にエラーが発生しました: {e}"
                            else:
                                # テキストファイル（csv/txt）なら今までのデコード処理
                                raw_data = file_obj.getvalue()
                                try:
                                    result = raw_data.decode("utf-8")
                                except UnicodeDecodeError:
                                    result = raw_data.decode("cp932")
                        else:
                            result = f"エラー：ファイル「{target_name}」が見つかりません。"
                    
                    else:
                        # save_to_new_excel など、他のツールは今まで通り関数を呼ぶ
                        fn = TOOL_MAP.get(name)
                        result = fn(**args) if fn else "ツールが見つかりません"
                    with st.expander(f"詳細ログ: {name}"):
                        st.write(f"**引数:** {args}")
                        st.write(f"**実行結果:**")
                        st.code(result)
                    st.session_state.messages.append({"role": "tool", "content": str(result)})