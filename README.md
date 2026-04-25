AIハイブリット企業評価ツール

AI × Rule-based Hybrid Company Evaluation Tool for Job Hunting

📌 Overview / 概要

このプロジェクトは、就職活動における企業研究を効率化するために開発した AI（ChatGPT API）による一次評価と 人間の補正ロジックを組み合わせた “ハイブリッド企業評価システム” です。

· Mode1（初期化）：企業評価用 Excel テンプレートを自動生成（最初の1回だけ使用）

· Mode2（通常運用）：企業名を入力するだけで AI が暫定評価を実行し、補正値を優先した最終評価を自動生成

Excel を中心に完結するため、操作がシンプルで扱いやすく、 AI の客観性と人間の判断を両立した企業評価が可能になります。

🧩 Mode1：Excel テンプレート自動生成（初期化）

Mode1 は 最初の1回だけ使う、または 台帳をリセットしたいときに使うモードです。

🔧 役割

· 評価に必要な16列を持つ Excel（data.xlsx）を自動生成

· ヘッダーのデザイン（背景色・太字・中央揃え）

· フィルタ設定、ウィンドウ枠固定

· Mode2 が読み込む標準フォーマットを保証

📁 出力される列

· 企業名

· 業種（AI / 補正）

· 安定性（AI / 補正）

· 年収（AI / 補正）

· 成長性（AI / 補正）

· WLB（AI / 補正）

· 総合点

· 判定

· コメント

· 更新日時

🧩 Mode1 コード例（抜粋）

python

import pandas as pd

from openpyxl import load_workbook

from openpyxl.styles import Font, PatternFill, Alignment


FILE_PATH = "data.xlsx"


COLUMNS = [

"企業名", "再評価フラグ",

"業種_AI", "業種_補正",

"安定性_AI", "安定性_補正",

"年収_AI", "年収_補正",

"成長性_AI", "成長性_補正",

"WLB_AI", "WLB_補正",

"総合点", "判定",

"企業評価コメント",

"システム実行コメント",

"更新日時"

]


def create_template_excel(file_path: str) -> None:

df = pd.DataFrame(columns=COLUMNS)

df.to_excel(file_path, index=False)


wb = load_workbook(file_path)

ws = wb.active


header_fill = PatternFill("solid", fgColor="D9EAF7")

header_font = Font(bold=True)

center = Alignment(horizontal="center", vertical="center")


for cell in ws[1]:

cell.fill = header_fill

cell.font = header_font

cell.alignment = center


ws.freeze_panes = "A2"

ws.auto_filter.ref = ws.dimensions


wb.save(file_path)

print(f"台帳を作り直しました: {file_path}")

🚀 Mode2：AI × 補正ロジックによる企業評価（通常運用）

Mode2 は 普段使うメイン機能です。

🧠 ① ChatGPT API による暫定評価

企業名を入力すると、自動で以下を生成します：

· 業種推定

· 安定性 / 年収 / 成長性 / WLB のスコア

· AI による企業評価コメント

✍️ ② 人間による補正（ユーザー入力）

補正値が入力されていれば 補正値を優先して最終スコアを決定。 コメントも 該当部分だけルールベースで置き換え、 AI の自然な文章は残す。

🧩 ③ ハイブリッド最終評価

· AI の自然な文章 ×

· 人間の判断を反映した補正コメント

これにより、 「AI の客観性」と「人間の判断」を両立した評価コメントが完成します。

🚀 Mode2 コード例（抜粋）

python

import openai

import pandas as pd

from datetime import datetime


def generate_ai_evaluation(company_name: str) -> dict:

prompt = f"""

以下の企業について、業種推定・安定性・年収・成長性・WLB の5項目を

10点満点で評価し、短いコメントを生成してください。


企業名: {company_name}

"""


response = openai.ChatCompletion.create(

model="gpt-4o-mini",

messages=[{"role": "user", "content": prompt}]

)


return response.choices[0].message["content"]


def apply_correction(ai_value: int, correction_value: int) -> int:

return correction_value if correction_value != "" else ai_value


def update_excel(file_path: str):

df = pd.read_excel(file_path)


for idx, row in df.iterrows():

if pd.isna(row["企業名"]):

continue


if row["再評価フラグ"] == 1:

ai_result = generate_ai_evaluation(row["企業名"])

# AI 結果をパースして DataFrame に反映（省略）


df.at[idx, "成長性_最終"] = apply_correction(

row["成長性_AI"], row["成長性_補正"]

)


df.at[idx, "更新日時"] = datetime.now()


df.to_excel(file_path, index=False)

🎯 特徴企業名を入力するだけで AI が暫定評価を生成

· 補正値を優先するルールベースロジック

· コメントは「補正された部分だけ」置き換え

· Excel ベースで完結するため、操作が簡単

· AI の文章の自然さと、人間の判断の正確さを両立

🧑‍💻 技術スタック

· Python 3.x

· pandas

· openpyxl

· ChatGPT API

· Excel（データ管理）
