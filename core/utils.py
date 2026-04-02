import requests
from datetime import datetime

API_URL = "https://script.google.com/macros/s/AKfycbyGv08iyugoWolQlg2AGZzZxooQy3nqd_S1x7n5GOTH0mwlqz-FpbldIuMPp-HJMwKI/exec?app_type=zabbix_traffic_robo"
APP_TOKEN = "ROBO_zabbixtrafficcapturer@ratiPRAY"

def get_internet_time():
    """Ambil waktu real dari internet (Asia/Jakarta) biar gak bisa dimanipulasi"""
    try:
        resp = requests.get("http://worldtimeapi.org/api/timezone/Asia/Jakarta", timeout=5)
        data = resp.json()
        return datetime.fromisoformat(data["datetime"].replace("Z", "+00:00"))
    except:
        return datetime.now()  # fallback

def check_license():
    """Cek token + expired ke API"""
    try:
        resp = requests.get(API_URL, timeout=10)
        data = resp.json()

        api_token = data.get("token")
        expired_str = data.get("expired")

        if api_token != APP_TOKEN:
            return False, "❌ Token tidak sesuai! Hubungi admin."

        expired = datetime.fromisoformat(expired_str.replace("Z", "+00:00"))
        current = get_internet_time()

        if current > expired:
            return False, f"❌ Automation sudah expired sejak {expired.strftime('%d %b %Y %H:%M')}"

        return True, "✅ License OK"
    except Exception as e:
        return False, f"❌ Gagal cek license: {e}"