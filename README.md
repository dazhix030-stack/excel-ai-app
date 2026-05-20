# Excel AI App

ローカルLLM（Qwen2.5-7B）を使ったExcel集計AIアプリ。自然言語で指示するだけでExcelの分析・集計・保存ができる。

## 機能

- Excel / CSV / TXT ファイルをアップロードして自然言語で指示
- ファイル形式を自動判別（xlsx / csv / txt）
- AIがデータを読み取り、集計・分析・保存まで対応
- 結果をオートフィルタ付きExcelで保存

## 使用技術

- Streamlit（UI）
- Ollama + Qwen2.5-7B（ローカルLLM）
- pandas / openpyxl（Excel処理）
- Function Calling（ツール呼び出し）

## セットアップ

```bash
pip install streamlit ollama pandas openpyxl
```

Ollamaのインストールと起動が必要です：

```bash
ollama pull qwen2.5:7b
ollama serve
```

## 実行方法

```bash
streamlit run app.py
```

## ファイル構成

```
app.py          # StreamlitのメインUI・LLM連携
write_tools.py  # Excel読み書きツール関数
```

## 動作環境

- Python 3.12
- Ollama（ローカルLLM実行環境）
- GPU推奨（CPU動作も可）
