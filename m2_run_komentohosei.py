import os
import json
from datetime import datetime
import re

import pandas as pd
from openai import OpenAI

# =========================
# 固定設定
# =========================
FILE_PATH = "data.xlsx"
BACKUP_FILE_PATH = "data_backup.xlsx"
OPENAI_API_ENV_NAME = "OPENAI_API_KEY"

WEIGHTS = {
    "安定性": 0.35,
    "年収": 0.20,
    "成長性": 0.20,
    "WLB": 0.25,
}

INDUSTRY_CANDIDATES = [
    "情報・通信",
    "金融・リース",
    "保険",
    "メーカー",
    "商社",
    "建設・不動産",
    "インフラ",
    "小売",
    "サービス",
    "物流・運輸",
    "官公庁・公的機関",
    "その他",
]

REQUIRED_COLUMNS = [
    "企業名",
    "再評価フラグ",
    "業種_AI",
    "業種_補正",
    "安定性_AI",
    "安定性_補正",
    "年収_AI",
    "年収_補正",
    "成長性_AI",
    "成長性_補正",
    "WLB_AI",
    "WLB_補正",
    "総合点",
    "判定",
    "企業評価コメント",
    "システム実行コメント",
    "更新日時"
]

REQUIRED_AI_COLS = ["業種_AI", "安定性_AI", "年収_AI", "成長性_AI", "WLB_AI"]


# =========================
# 共通関数
# =========================
def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_empty(value) -> bool:
    return normalize_text(value) == ""


def is_reval_requested(value) -> bool:
    text = normalize_text(value).upper()
    return text in ["1", "Y", "YES", "TRUE"]


def pick_final_value(ai_value, manual_value):
    return manual_value if not is_empty(manual_value) else ai_value


def safe_score(value):
    if is_empty(value):
        return None
    try:
        n = int(float(value))
        return max(1, min(5, n))
    except Exception:
        return None


def calc_total_score(stability, salary, growth, wlb):
    vals = [stability, salary, growth, wlb]
    if any(v is None for v in vals):
        return None
    total = (
        stability * WEIGHTS["安定性"] +
        salary    * WEIGHTS["年収"] +
        growth    * WEIGHTS["成長性"] +
        wlb       * WEIGHTS["WLB"]
    )
    return round(total, 2)


def judge_rank(total_score):
    if total_score is None:
        return ""
    if total_score >= 4.2:
        return "S"
    if total_score >= 3.6:
        return "A"
    if total_score >= 3.0:
        return "B"
    if total_score >= 2.4:
        return "C"
    return "D"


def row_needs_ai_fetch(row: pd.Series) -> bool:
    if is_reval_requested(row.get("再評価フラグ", "")):
        return True
    for col in REQUIRED_AI_COLS:
        if is_empty(row.get(col, "")):
            return True
    return False


# =========================
# OpenAI
# =========================
def get_client() -> OpenAI:
    api_key = os.environ.get(OPENAI_API_ENV_NAME)
    if not api_key:
        raise RuntimeError(
            f"環境変数 {OPENAI_API_ENV_NAME} が見つかりません。"
        )
    return OpenAI(api_key=api_key)


def infer_industry(client: OpenAI, company_name: str) -> str:
    prompt = f"""
あなたは日本企業の業種分類アシスタントです。
次の企業の主たる業種を、必ず候補から1つだけ選んでください。

候補:
{chr(10).join(f"- {x}" for x in INDUSTRY_CANDIDATES)}

企業名: {company_name}

出力ルール:
- 業種名だけを返す
- 説明不要
"""
    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=prompt
        )
        answer = response.output_text.strip()
        return answer if answer in INDUSTRY_CANDIDATES else "その他"
    except Exception:
        return ""


def evaluate_company_basic(client: OpenAI, company_name: str, industry_name: str) -> dict:
    prompt = f"""
あなたは日本の就活企業評価アシスタントです。
次の企業について暫定評価してください。

企業名: {company_name}
業種: {industry_name}

評価項目:
- 安定性
- 年収
- 成長性
- WLB

ルール:
- 各項目は1〜5の整数
- 企業評価コメントは80字以内
- 必ずJSONだけを返す
- 情報が弱くても止まらないように、可能な範囲で暫定値を返す

JSON形式:
{{
  "安定性_AI": 4,
  "年収_AI": 3,
  "成長性_AI": 4,
  "WLB_AI": 3,
  "企業評価コメント": "短いコメント"
}}
"""
    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=prompt
        )
        return json.loads(response.output_text.strip())
    except Exception:
        return {
            "安定性_AI": "",
            "年収_AI": "",
            "成長性_AI": "",
            "WLB_AI": "",
            "企業評価コメント": ""
        }
# =========================
# 業種別キーワード辞書
# =========================
INDUSTRY_KEYWORDS = {
    "情報・通信": {
        "stability": ["技術変化リスク", "競争激化"],
        "salary": ["高水準の報酬", "成果連動"],
        "growth": ["DX需要", "技術革新", "市場拡大"],
        "wlb": ["リモートワーク", "柔軟な働き方"],
    },
    "金融・リース": {
        "stability": ["規制環境", "景気影響"],
        "salary": ["高い給与水準", "資格手当"],
        "growth": ["フィンテック", "金融商品多様化"],
        "wlb": ["繁忙期の負荷", "安定した働き方"],
    },
    "メーカー": {
        "stability": ["供給網リスク", "生産体制"],
        "salary": ["年功序列傾向", "安定した給与"],
        "growth": ["海外展開", "製品開発力"],
        "wlb": ["現場負荷", "休日取得"],
    },
    # 必要に応じて追加
}


# =========================
# コメント分割・分類
# =========================
def split_sentences(text: str):
    sentences = re.split(r"[。]", text)
    return [s.strip() for s in sentences if s.strip()]


def classify_sentence(sentence: str):
    if any(k in sentence for k in ["安定", "基盤", "倒産", "リスク"]):
        return "安定性"
    if any(k in sentence for k in ["年収", "給与", "報酬"]):
        return "年収"
    if any(k in sentence for k in ["成長", "将来性", "伸び", "市場"]):
        return "成長性"
    if any(k in sentence for k in ["WLB", "働き", "残業", "休暇"]):
        return "WLB"
    return "その他"


# =========================
# 業種別ルールベースコメント生成
# =========================
def build_stability_comment(score, industry):
    words = INDUSTRY_KEYWORDS.get(industry, {}).get("stability", [])
    w = "、".join(words)

    if score >= 5:
        return f"安定性は非常に高く、{w}といった点でも強みが見られます。"
    if score >= 4:
        return f"安定性は高く、{w}が支えとなっています。"
    if score >= 3:
        return f"安定性は標準的で、{w}の影響を受ける可能性があります。"
    if score >= 2:
        return f"安定性にはやや懸念があり、{w}への注意が必要です。"
    return f"安定性は低めで、{w}がリスク要因となり得ます。"


def build_salary_comment(score, industry):
    words = INDUSTRY_KEYWORDS.get(industry, {}).get("salary", [])
    w = "、".join(words)

    if score >= 5:
        return f"年収は非常に高く、{w}といった点でも魅力があります。"
    if score >= 4:
        return f"年収は比較的高く、{w}が期待できます。"
    if score >= 3:
        return f"年収は標準的で、{w}の傾向があります。"
    if score >= 2:
        return f"年収はやや低めで、{w}の影響が考えられます。"
    return f"年収は低めで、{w}の面で課題が見られます。"


def build_growth_comment(score, industry):
    words = INDUSTRY_KEYWORDS.get(industry, {}).get("growth", [])
    w = "、".join(words)

    if score >= 5:
        return f"成長性は非常に高く、{w}が強みとして期待できます。"
    if score >= 4:
        return f"成長性は高く、{w}が追い風となっています。"
    if score >= 3:
        return f"成長性は標準的で、今後の{w}の動向が鍵となります。"
    if score >= 2:
        return f"成長性にはやや懸念があり、{w}の影響を注視する必要があります。"
    return f"成長性は低めで、{w}の面で課題が見られます。"


def build_wlb_comment(score, industry):
    words = INDUSTRY_KEYWORDS.get(industry, {}).get("wlb", [])
    w = "、".join(words)

    if score >= 5:
        return f"WLBは非常に良好で、{w}が実現されています。"
    if score >= 4:
        return f"WLBは良好で、{w}が整っています。"
    if score >= 3:
        return f"WLBは標準的で、{w}が影響する可能性があります。"
    if score >= 2:
        return f"WLBにはやや課題があり、{w}への配慮が必要です。"
    return f"WLBは低めで、{w}の面で負荷が懸念されます。"


# =========================
# AIコメントの該当部分だけ置換
# =========================
def patch_comment(ai_comment, final_scores, industry, row):
    sentences = split_sentences(ai_comment)
    patched = []

    for s in sentences:
        category = classify_sentence(s)

        if category == "安定性" and not is_empty(row["安定性_補正"]):
            patched.append(build_stability_comment(final_scores["安定性"], industry))
            continue

        if category == "年収" and not is_empty(row["年収_補正"]):
            patched.append(build_salary_comment(final_scores["年収"], industry))
            continue

        if category == "成長性" and not is_empty(row["成長性_補正"]):
            patched.append(build_growth_comment(final_scores["成長性"], industry))
            continue

        if category == "WLB" and not is_empty(row["WLB_補正"]):
            patched.append(build_wlb_comment(final_scores["WLB"], industry))
            continue

        patched.append(s)

    return "。".join(patched) + "。"
# =========================
# 補正が入っているか判定
# =========================
def has_manual_correction(row: pd.Series) -> bool:
    correction_cols = ["安定性_補正", "年収_補正", "成長性_補正", "WLB_補正"]
    return any(not is_empty(row.get(col, "")) for col in correction_cols)


# =========================
# 行処理（AI＋補正時のみパッチ）
# =========================
def process_row(client: OpenAI, row: pd.Series) -> pd.Series:
    company_name = normalize_text(row.get("企業名", ""))

    if company_name == "":
        row["システム実行コメント"] = "企業名空欄のため未処理"
        return row

    system_logs = []

    # 1. AI再取得要否判定
    needs_ai = row_needs_ai_fetch(row)

    # 2. 必要ならAI取得
    if needs_ai:
        system_logs.append("AI再取得対象")

        # 業種取得
        if is_empty(row.get("業種_AI", "")) or is_reval_requested(row.get("再評価フラグ", "")):
            industry_ai = infer_industry(client, company_name)
            if not is_empty(industry_ai):
                row["業種_AI"] = industry_ai
                system_logs.append("業種_AI更新")
            else:
                system_logs.append("業種_AI取得失敗")

        industry_for_eval = pick_final_value(row.get("業種_AI", ""), row.get("業種_補正", ""))

        # 評価取得
        result = evaluate_company_basic(client, company_name, normalize_text(industry_for_eval))

        for col in ["安定性_AI", "年収_AI", "成長性_AI", "WLB_AI"]:
            val = result.get(col, "")
            if not is_empty(val):
                row[col] = val

        comment = normalize_text(result.get("企業評価コメント", ""))
        if comment != "":
            row["企業評価コメント"] = comment
            system_logs.append("AI企業評価コメント更新")
        else:
            system_logs.append("企業評価コメント未取得")

    else:
        system_logs.append("AI再取得スキップ")

    # 3. 補正優先で毎回再計算
    final_stability = safe_score(pick_final_value(row.get("安定性_AI", ""), row.get("安定性_補正", "")))
    final_salary    = safe_score(pick_final_value(row.get("年収_AI", ""), row.get("年収_補正", "")))
    final_growth    = safe_score(pick_final_value(row.get("成長性_AI", ""), row.get("成長性_補正", "")))
    final_wlb       = safe_score(pick_final_value(row.get("WLB_AI", ""), row.get("WLB_補正", "")))

    total_score = calc_total_score(final_stability, final_salary, final_growth, final_wlb)

    row["総合点"] = total_score if total_score is not None else ""
    row["判定"] = judge_rank(total_score)

    if total_score is None:
        system_logs.append("必要指標不足のため総合点未計算")
    else:
        system_logs.append("総合点・判定を再計算")

    # 4. コメント処理（補正が入ったときだけパッチ適用）
    if has_manual_correction(row):
        industry = pick_final_value(row.get("業種_AI", ""), row.get("業種_補正", ""))

        row["企業評価コメント"] = patch_comment(
            row.get("企業評価コメント", ""),
            {
                "安定性": final_stability,
                "年収": final_salary,
                "成長性": final_growth,
                "WLB": final_wlb,
            },
            industry,
            row
        )
        system_logs.append("補正あり → コメント部分置換（業種別パッチ適用）")
    else:
        system_logs.append("補正なし → AIコメントを維持")

    # 5. 再評価フラグは使ったら消す
    if is_reval_requested(row.get("再評価フラグ", "")):
        row["再評価フラグ"] = ""
        system_logs.append("再評価フラグをクリア")

    row["更新日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row["システム実行コメント"] = " / ".join(system_logs)

    return row


# =========================
# 保存
# =========================
def safe_save_excel(df: pd.DataFrame, file_path: str):
    try:
        df.to_excel(file_path, index=False)
        print(f"上書き保存しました: {file_path}")
    except PermissionError:
        df.to_excel(BACKUP_FILE_PATH, index=False)
        print("元ファイルへ保存できませんでした。Excelを閉じているか確認してください。")
        print(f"退避保存しました: {BACKUP_FILE_PATH}")


# =========================
# メイン
# =========================
def update_master(file_path: str) -> None:
    if not os.path.exists(file_path):
        print(f"ファイルが見つかりません: {file_path}")
        print("先にモード1で data.xlsx を作成してください。")
        return

    try:
        df = pd.read_excel(file_path)
    except PermissionError:
        print("Excelファイルが開いている可能性があります。保存して閉じてから再実行してください。")
        return
    except Exception as e:
        print(f"Excel読込失敗: {e}")
        return

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        print("Excelの列構造が不足しています。")
        print("不足列:", ", ".join(missing_cols))
        print("モード1で正しいテンプレートを再作成してください。")
        return

    try:
        client = get_client()
    except Exception as e:
        print(f"API初期化失敗: {e}")
        return

    processed_rows = []
    for _, row in df.iterrows():
        try:
            processed_rows.append(process_row(client, row))
        except Exception as e:
            row["システム実行コメント"] = f"行処理失敗: {e}"
            processed_rows.append(row)

    out_df = pd.DataFrame(processed_rows)
    safe_save_excel(out_df, file_path)


if __name__ == "__main__":
    update_master(FILE_PATH)
