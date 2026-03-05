import os
import json
import google.genai as genai
from google.genai import types

api_key = os.environ.get("GEMINI_API_KEY")
if not api_key:
    # try to get from secrets.toml
    import toml
    try:
        with open("c:\\Users\\HP\\OneDrive\\ドキュメント\\bpo\\.streamlit\\secrets.toml", "r", encoding="utf-8") as f:
            secrets = toml.load(f)
            api_key = secrets["GEMINI_API_KEY"]
    except Exception:
        pass

client = genai.Client(api_key=api_key)

recent_list_text = """[行番号: 10] 自治体: 港区, 案件名: 放課後児童クラブ運営業務, 年度: 令和7年度
[行番号: 11] 自治体: 新宿区, 案件名: 学童クラブ委託, 年度: 2025年度
[行番号: 12] 自治体: 渋谷区, 案件名: 子どもセンター運営, 年度: 令和6年度
"""

new_groups_text = """[新規グループ 0] 自治体: 東京都港区, 案件名: 放課後児童クラブ運営業務委託, 年度: 2025年度
[新規グループ 1] 自治体: 新宿区, 案件名: 学童クラブ運営業務委託, 年度: 令和7年度
[新規グループ 2] 自治体: 世田谷区, 案件名: 新規案件, 年度: 令和7年度
"""

MATCH_SCHEMA = {
    "type": "object",
    "properties": {
        "matches": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "group_index": {"type": "integer"},
                    "matched_existing_row": {"type": "integer"}
                }
            }
        }
    }
}

prompt = f"""以下は、スプレッドシートに登録済みの直近20件の案件データと、今回新しく読み取った案件のリストです。
新しく読み取った案件が、既存のデータと「同一案件」であるかを確認してください。

【判定ルール】
1. 和暦と西暦の違い（例：令和7年度 と 2025年度）は同一年度とみなす。
2. 「業務委託」「プロポーザル」「運営業務」といった末尾の単語の有無などの表記揺れは、文脈から判断して実質的に同じ案件であれば同一とみなす。
3. 一致する既存案件が見つかった場合は、その「行番号」を返してください。
4. 一致しない場合は、その新規グループについてはリストに含めない、または null として処理してください。

▼ 既存のデータ（直近20件）
{recent_list_text}

▼ 今回新しく読み取った案件
{new_groups_text}
"""

response = client.models.generate_content(
    model="gemini-3-flash-preview",
    contents=[prompt],
    config=types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=MATCH_SCHEMA,
        temperature=0.1,
    )
)
print(response.text)
