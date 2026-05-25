"""
포트폴리오 원그래프 웹서버
실행: python3 portfolio_server.py
브라우저: http://localhost:5050
"""
import json
import os
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs

# yfinance 동적 임포트 (없으면 pip install 안내)
try:
    import yfinance as yf
    HAS_YFINANCE = True
except ImportError:
    HAS_YFINANCE = False


PORT = 5050
HTML_FILE = os.path.join(os.path.dirname(__file__), "portfolio_chart.html")


def fetch_price(ticker: str):
    """ticker 주가를 반환. 실패 시 (None, 에러메시지)."""
    if not HAS_YFINANCE:
        return None, "yfinance 미설치 — pip install yfinance 실행 후 재시작"
    try:
        t = yf.Ticker(ticker)
        # fast_info 먼저 시도
        try:
            info = t.fast_info
            price = getattr(info, "last_price", None) or getattr(info, "previous_close", None)
            if price and float(price) > 0:
                return round(float(price), 4), None
        except Exception:
            pass
        # fallback: history
        try:
            hist = t.history(period="2d")
            if not hist.empty:
                return round(float(hist["Close"].iloc[-1]), 4), None
        except Exception:
            pass
        return None, f"'{ticker}' 가격 데이터를 찾을 수 없습니다 (티커를 확인해 주세요)"
    except Exception as e:
        msg = str(e)
        if "403" in msg or "allowlist" in msg.lower():
            return None, "네트워크 차단: Yahoo Finance에 접근할 수 없습니다"
        return None, f"조회 실패: {msg[:80]}"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # 터미널 로그 간소화
        print(f"  {self.address_string()} {args[0]} {args[1]}")

    def send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")

        # ── API: 주가 조회 ──────────────────────────────────────
        if path.startswith("/api/price/"):
            ticker = path.split("/api/price/", 1)[-1].strip().upper()
            if not ticker:
                self.send_json(400, {"error": "티커 없음"})
                return
            price, err = fetch_price(ticker)
            if price is not None:
                self.send_json(200, {"ticker": ticker, "price": price})
            else:
                self.send_json(404, {"error": err or "조회 실패"})
            return

        # ── 메인 HTML ───────────────────────────────────────────
        if path in ("", "/"):
            try:
                with open(HTML_FILE, "rb") as f:
                    body = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", len(body))
                self.end_headers()
                self.wfile.write(body)
            except FileNotFoundError:
                self.send_json(404, {"error": "portfolio_chart.html 파일 없음"})
            return

        self.send_json(404, {"error": "Not found"})


def main():
    if not HAS_YFINANCE:
        print("⚠  yfinance가 없습니다. 설치 중...")
        os.system(f"{sys.executable} -m pip install yfinance --quiet")
        try:
            import yfinance  # noqa: F401
            print("✓ yfinance 설치 완료")
        except ImportError:
            print("✗ yfinance 설치 실패. 주가 자동 조회가 비활성화됩니다.")

    server = HTTPServer(("localhost", PORT), Handler)
    url = f"http://localhost:{PORT}"
    print(f"\n포트폴리오 서버 시작: {url}")
    print("종료: Ctrl+C\n")

    # 1초 후 브라우저 자동 오픈
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버 종료.")


if __name__ == "__main__":
    main()
