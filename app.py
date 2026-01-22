"""
終電検索 LINE Bot

LINEで駅名を送ると、その駅から西千葉駅への終電を調べて返信する
"""

import os
import re
import json
import requests
from urllib.parse import quote
from datetime import datetime, timezone, timedelta
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent
from linebot.v3.exceptions import InvalidSignatureError
from bs4 import BeautifulSoup

app = Flask(__name__)

# 環境変数から設定を取得
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")

# LINE Bot SDK v3 の設定
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 目的地
DESTINATION_STATION = "西千葉"

# Yahoo!路線情報（関東エリア）の遅延情報URL
YAHOO_TRANSIT_URL = "https://transit.yahoo.co.jp/diainfo/area/3"

# 西千葉駅に関連する路線（遅延チェック対象）
RELATED_LINES = [
    "総武線",
    "総武本線",
    "中央・総武線",
    "京葉線",
    "武蔵野線",
    "京成線",
    "京成千葉線",
]


@app.route("/")
def index():
    """ヘルスチェック用"""
    return "LINE Last Train Bot is running!"


@app.route("/callback", methods=["POST"])
def callback():
    """LINE Webhookのエンドポイント"""
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)

    app.logger.info(f"Request body: {body}")

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        app.logger.error("Invalid signature")
        abort(400)

    return "OK"


@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    """テキストメッセージを受信したときの処理"""
    user_message = event.message.text.strip()
    app.logger.info(f"Received message: {user_message}")

    # 駅名を抽出（シンプルにメッセージ全体を駅名として扱う）
    station_name = extract_station_name(user_message)

    if not station_name:
        reply_text = "駅名を入力してください。\n例: 東京、秋葉原、千葉"
    else:
        # 終電を検索
        reply_text = search_last_train(station_name, DESTINATION_STATION)

    # 返信
    with ApiClient(configuration) as api_client:
        messaging_api = MessagingApi(api_client)
        messaging_api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=reply_text)]
            )
        )


def extract_station_name(message):
    """
    メッセージから駅名を抽出する

    Args:
        message: ユーザーのメッセージ

    Returns:
        str: 抽出した駅名（見つからなければNone）
    """
    # 「〜駅」という形式があれば抽出
    match = re.search(r"(.+?)駅", message)
    if match:
        return match.group(1)

    # 「〜から」「〜で」などのパターン
    patterns = [
        r"(.+?)から(?:帰|終電|電車)",
        r"(.+?)で(?:飲|遊|仕事)",
        r"^([ぁ-んァ-ン一-龥a-zA-Z]+)$",  # 駅名のみ
    ]

    for pattern in patterns:
        match = re.search(pattern, message)
        if match:
            station = match.group(1).strip()
            # 短すぎる場合は除外
            if len(station) >= 1:
                return station

    # パターンにマッチしなければメッセージ全体を駅名として扱う
    # ただし長すぎる場合は除外
    if len(message) <= 10:
        return message

    return None


def search_last_train(from_station, to_station):
    """
    Yahoo!路線情報で終電を検索する

    Args:
        from_station: 出発駅
        to_station: 到着駅

    Returns:
        str: 検索結果のメッセージ
    """
    try:
        # 日本時間で今日の日付を取得
        jst = timezone(timedelta(hours=9))
        now = datetime.now(jst)

        # 遅延情報を取得
        delays = fetch_delay_info()

        # Yahoo!路線情報のURL
        # type=1: 出発時刻指定、深夜23:50で検索して終電を取得
        url = (
            f"https://transit.yahoo.co.jp/search/result"
            f"?from={quote(from_station)}"
            f"&to={quote(to_station)}"
            f"&y={now.year}&m={now.month:02d}&d={now.day:02d}"
            f"&hh=23&mm=50"
            f"&type=1"  # 出発時刻指定
            f"&ticket=ic"  # IC優先
        )

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }

        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        response.encoding = "utf-8"

        soup = BeautifulSoup(response.text, "html.parser")

        # エラーチェック（駅が見つからない場合）
        error_elem = soup.select_one("div.elmErrorText, p.errTxt")
        if error_elem:
            return f"「{from_station}」が見つかりませんでした。\n正式な駅名で入力してください。"

        # 候補駅が複数ある場合
        candidate_list = soup.select("div.candiList a, ul.candidate a")
        if candidate_list:
            candidates = [a.get_text(strip=True) for a in candidate_list[:5]]
            return f"「{from_station}」に該当する駅が複数あります:\n" + "\n".join(f"・{c}" for c in candidates)

        # JSONデータから経路情報を取得
        result = parse_route_from_json(soup, from_station, to_station)

        if not result:
            # JSONが取得できない場合はHTMLからパース
            result = parse_route_result(soup, from_station, to_station)

        if not result:
            return f"終電情報を取得できませんでした。\nYahoo!路線情報で直接検索してください。"

        # 遅延情報があれば追加
        if delays:
            delay_msg = format_delay_info(delays)
            result = result + "\n" + delay_msg

            # 遅延時は代替ルートを検索
            alt_result = search_alternative_route(from_station, to_station, now)
            if alt_result:
                result = result + "\n" + alt_result

        return result

    except requests.Timeout:
        return "検索がタイムアウトしました。\nしばらくしてからもう一度お試しください。"
    except requests.RequestException as e:
        app.logger.error(f"Request error: {e}")
        return "検索中にエラーが発生しました。\nしばらくしてからもう一度お試しください。"


def format_delay_info(delays):
    """
    遅延情報をフォーマットする

    Args:
        delays: 遅延情報のリスト

    Returns:
        str: フォーマットされた遅延情報
    """
    lines = [
        "",
        "⚠️ 運行情報 ⚠️",
    ]

    for delay in delays[:3]:  # 最大3件
        lines.append(f"🔴 {delay['line']}")
        lines.append(f"   {delay['status'][:50]}")

    return "\n".join(lines)


def search_alternative_route(from_station, to_station, now):
    """
    代替ルートを検索する（遅延時用）

    Args:
        from_station: 出発駅
        to_station: 到着駅
        now: 現在時刻

    Returns:
        str: 代替ルート情報（なければNone）
    """
    try:
        # 経由駅を変えて検索（例：東京経由、船橋経由など）
        via_stations = ["船橋", "津田沼", "千葉"]

        for via in via_stations:
            if via == from_station or via == to_station:
                continue

            url = (
                f"https://transit.yahoo.co.jp/search/result"
                f"?from={quote(from_station)}"
                f"&to={quote(to_station)}"
                f"&via={quote(via)}"
                f"&y={now.year}&m={now.month:02d}&d={now.day:02d}"
                f"&hh=23&mm=30"
                f"&type=1"
                f"&ticket=ic"
            )

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            }

            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")

                # 時刻を取得
                time_match = re.search(r"(\d{1,2}:\d{2})発.*?(\d{1,2}:\d{2})着", soup.get_text())
                if time_match:
                    return f"\n💡 代替ルート（{via}経由）\n   発車 {time_match.group(1)} → 到着 {time_match.group(2)}"

    except Exception as e:
        app.logger.error(f"Alternative route search error: {e}")

    return None


def fetch_delay_info():
    """
    Yahoo!路線情報から関連路線の遅延情報を取得する

    Returns:
        list: 遅延情報のリスト [{"line": 路線名, "status": 状況}]
    """
    delays = []

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(YAHOO_TRANSIT_URL, headers=headers, timeout=10)
        response.raise_for_status()
        response.encoding = response.apparent_encoding

        soup = BeautifulSoup(response.text, "html.parser")

        # 運行情報のリストを取得
        trouble_items = soup.select("div.trouble a, li.elmTblLstLine a")

        for item in trouble_items:
            line_name = item.get_text(strip=True)

            # 関連路線かチェック
            for target in RELATED_LINES:
                if target in line_name:
                    # 詳細ページのURLを取得
                    detail_url = item.get("href")
                    if detail_url and not detail_url.startswith("http"):
                        detail_url = "https://transit.yahoo.co.jp" + detail_url

                    status = fetch_delay_detail(detail_url) if detail_url else "運行情報あり"

                    delays.append({
                        "line": line_name,
                        "status": status
                    })
                    break

    except requests.RequestException as e:
        app.logger.error(f"遅延情報の取得に失敗: {e}")

    return delays


def fetch_delay_detail(url):
    """
    詳細ページから遅延状況を取得する

    Args:
        url: 詳細ページのURL

    Returns:
        str: 遅延状況の詳細
    """
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        response.encoding = response.apparent_encoding

        soup = BeautifulSoup(response.text, "html.parser")

        # 状況の詳細を取得
        status_elem = soup.select_one("div.trouble p, p.trouble, div.statusTxt")
        if status_elem:
            return status_elem.get_text(strip=True)[:100]  # 100文字まで

    except requests.RequestException:
        pass

    return "詳細情報取得失敗"


def parse_route_from_json(soup, from_station, to_station):
    """
    ページ内のJSONデータから経路情報を取得する

    Args:
        soup: BeautifulSoupオブジェクト
        from_station: 出発駅
        to_station: 到着駅

    Returns:
        str: フォーマットされた結果メッセージ
    """
    try:
        # scriptタグからJSONデータを探す
        scripts = soup.find_all("script")
        route_data = None

        for script in scripts:
            if script.string and "featureInfoList" in script.string:
                # naviSearchParam を探す
                match = re.search(r'naviSearchParam\s*=\s*(\{.+?\});', script.string, re.DOTALL)
                if match:
                    try:
                        route_data = json.loads(match.group(1))
                        break
                    except json.JSONDecodeError:
                        continue

        if not route_data:
            return None

        # 最初のルート（終電に最も近い）を取得
        feature_list = route_data.get("featureInfoList", [])
        edge_list = route_data.get("edgeInfoList", [])

        if not feature_list or not edge_list:
            return None

        # 最初のルートの情報
        feature = feature_list[0]
        edges = edge_list[0] if edge_list else []

        # 発車・到着時刻を取得（HH:MM形式のみ）
        dep_time = feature.get("departureTime", "")
        arr_time = feature.get("arrivalTime", "")

        # 時刻からHH:MM部分のみ抽出
        time_match = re.search(r"(\d{1,2}:\d{2})", dep_time)
        if time_match:
            dep_time = time_match.group(1)

        time_match = re.search(r"(\d{1,2}:\d{2})", arr_time)
        if time_match:
            arr_time = time_match.group(1)

        # 乗換回数
        transfer_count = feature.get("transferCount", 0)

        # 路線情報と番線を取得
        route_details = []
        for edge in edges:
            if isinstance(edge, dict):
                rail_name = edge.get("railName", "")
                # 番線情報
                riding_info = edge.get("ridingPositionInfo", {})
                if riding_info:
                    dep_platform = riding_info.get("departure", "")
                    if rail_name and dep_platform:
                        route_details.append(f"{rail_name}（{dep_platform}）")
                    elif rail_name:
                        route_details.append(rail_name)
                elif rail_name:
                    route_details.append(rail_name)

        # 結果をフォーマット
        lines = [
            f"🚃 {from_station} → {to_station}",
            "",
            f"発車 {dep_time}",
            f"到着 {arr_time}",
        ]

        if transfer_count > 0:
            lines.append(f"乗換 {transfer_count}回")

        if route_details:
            lines.append("")
            for detail in route_details[:3]:  # 最大3路線
                lines.append(f"▶ {detail}")

        lines.extend([
            "",
            "※ 運行状況により変更の場合あり",
        ])

        return "\n".join(lines)

    except Exception as e:
        app.logger.error(f"JSON parse error: {e}")
        return None


def parse_route_result(soup, from_station, to_station):
    """
    検索結果ページのHTMLから終電情報をパースする（フォールバック用）

    Args:
        soup: BeautifulSoupオブジェクト
        from_station: 出発駅
        to_station: 到着駅

    Returns:
        str: フォーマットされた結果メッセージ
    """
    try:
        # 時刻のパターン（HH:MM形式）を探す
        time_pattern = re.compile(r"(\d{1,2}:\d{2})発")
        arr_pattern = re.compile(r"(\d{1,2}:\d{2})着")

        all_text = soup.get_text()

        dep_match = time_pattern.search(all_text)
        arr_match = arr_pattern.search(all_text)

        dep_time = dep_match.group(1) if dep_match else None
        arr_time = arr_match.group(1) if arr_match else None

        if not dep_time:
            # 別のパターンを試す
            times = re.findall(r"(\d{1,2}:\d{2})", all_text)
            if len(times) >= 2:
                dep_time = times[0]
                arr_time = times[1]

        # 結果をフォーマット
        if dep_time:
            lines = [
                f"🚃 {from_station} → {to_station}",
                "",
                f"発車 {dep_time}",
            ]

            if arr_time:
                lines.append(f"到着 {arr_time}")

            lines.extend([
                "",
                "※ 詳細はYahoo!路線情報で確認してください",
            ])

            return "\n".join(lines)

        return None

    except Exception as e:
        app.logger.error(f"Parse error: {e}")
        return None


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
