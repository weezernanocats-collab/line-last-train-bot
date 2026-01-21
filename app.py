"""
終電検索 LINE Bot

LINEで駅名を送ると、その駅から西千葉駅への終電を調べて返信する
"""

import os
import re
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

        # Yahoo!路線情報のURL
        # type=4 は終電検索
        url = (
            f"https://transit.yahoo.co.jp/search/result"
            f"?from={quote(from_station)}"
            f"&to={quote(to_station)}"
            f"&type=4"  # 終電
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

        # 検索結果を取得
        route_elem = soup.select_one("div.routeList, ul.routeList")
        if not route_elem:
            # 別のセレクタを試す
            route_elem = soup.select_one("div#srline, div.searchResult")

        if not route_elem:
            return f"「{from_station}」→「{to_station}」の経路が見つかりませんでした。"

        # 時刻を取得
        result = parse_route_result(soup, from_station, to_station)

        if result:
            return result
        else:
            return f"終電情報を取得できませんでした。\nYahoo!路線情報で直接検索してください。"

    except requests.Timeout:
        return "検索がタイムアウトしました。\nしばらくしてからもう一度お試しください。"
    except requests.RequestException as e:
        app.logger.error(f"Request error: {e}")
        return "検索中にエラーが発生しました。\nしばらくしてからもう一度お試しください。"


def parse_route_result(soup, from_station, to_station):
    """
    検索結果ページから終電情報をパースする

    Args:
        soup: BeautifulSoupオブジェクト
        from_station: 出発駅
        to_station: 到着駅

    Returns:
        str: フォーマットされた結果メッセージ
    """
    try:
        # 出発時刻を取得
        dep_time = None
        arr_time = None

        # パターン1: li.time 内の時刻
        time_elems = soup.select("li.time")
        if len(time_elems) >= 2:
            dep_time = time_elems[0].get_text(strip=True)
            arr_time = time_elems[1].get_text(strip=True)

        # パターン2: span.departure, span.arrival
        if not dep_time:
            dep_elem = soup.select_one("span.departure, div.departure")
            arr_elem = soup.select_one("span.arrival, div.arrival")
            if dep_elem:
                dep_time = dep_elem.get_text(strip=True)
            if arr_elem:
                arr_time = arr_elem.get_text(strip=True)

        # パターン3: より汎用的な検索
        if not dep_time:
            # 時刻のパターン（HH:MM形式）を探す
            time_pattern = re.compile(r"\d{1,2}:\d{2}")
            all_text = soup.get_text()
            times = time_pattern.findall(all_text)
            if len(times) >= 2:
                dep_time = times[0]
                arr_time = times[1]

        # 所要時間を取得
        duration = None
        duration_elem = soup.select_one("li.requredTime, span.time, div.totalTime")
        if duration_elem:
            duration = duration_elem.get_text(strip=True)

        # 乗換回数を取得
        transfer = None
        transfer_elem = soup.select_one("li.transfer, span.transfer")
        if transfer_elem:
            transfer = transfer_elem.get_text(strip=True)

        # 路線名を取得
        line_names = []
        line_elems = soup.select("li.transport span, div.transport, span.lineName")
        for elem in line_elems[:3]:  # 最大3つまで
            line_name = elem.get_text(strip=True)
            if line_name and "円" not in line_name:
                line_names.append(line_name)

        # 結果をフォーマット
        if dep_time:
            lines = [
                f"🚃 {from_station} → {to_station} 終電",
                "",
                f"🕐 発車: {dep_time}",
            ]

            if arr_time:
                lines.append(f"🏁 到着: {arr_time}")

            if duration:
                lines.append(f"⏱️ 所要: {duration}")

            if transfer:
                lines.append(f"🔄 乗換: {transfer}")

            if line_names:
                lines.append(f"🚈 路線: {', '.join(line_names[:2])}")

            lines.extend([
                "",
                "※ 終電情報は変更される場合があります",
            ])

            return "\n".join(lines)

        return None

    except Exception as e:
        app.logger.error(f"Parse error: {e}")
        return None


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
