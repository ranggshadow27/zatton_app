import ntplib
import socket
from datetime import datetime, timezone, timedelta

API_URL = "https://script.google.com/macros/s/AKfycbyGv08iyugoWolQlg2AGZzZxooQy3nqd_S1x7n5GOTH0mwlqz-FpbldIuMPp-HJMwKI/exec?app_type=zabbix_traffic_robo"
APP_TOKEN = "ROBO_zabbixtrafficcapturdder@ratiPRAY"

# NTP Servers yang reliable
NTP_SERVERS = [
    'pool.ntp.org',
    'asia.pool.ntp.org',
    'id.pool.ntp.org',
    'time.google.com'
]

def get_internet_time():
    """Ambil waktu akurat dari NTP (WIB / Asia/Jakarta)"""
    client = ntplib.NTPClient()
    
    for server in NTP_SERVERS:
        try:
            response = client.request(server, version=3, timeout=5)
            # response.tx_time adalah waktu UTC
            utc_time = datetime.fromtimestamp(response.tx_time, tz=timezone.utc)
            # Tambah 7 jam untuk WIB
            jakarta_time = utc_time + timedelta(hours=7)
            return jakarta_time
        except (ntplib.NTPException, socket.timeout, OSError, Exception):
            continue  # coba server berikutnya
    
    # Fallback kalau semua NTP gagal
    print("Warning: NTP gagal, pakai system time sebagai fallback")
    return datetime.now() + timedelta(hours=7)   # anggap system time di WIB


def check_license():
    """Cek license pakai NTP time"""
    try:
        # Ganti dengan API lu
        import requests
        API_URL = "https://script.google.com/macros/s/AKfycbyGv08iyugoWolQlg2AGZzZxooQy3nqd_S1x7n5GOTH0mwlqz-FpbldIuMPp-HJMwKI/exec?app_type=zabbix_traffic_robo"
        APP_TOKEN = "ROBO_zabbixtrafficcapturer@ratiPRAY"

        resp = requests.get(API_URL, timeout=10)
        data = resp.json()

        api_token = data.get("token")
        expired_str = data.get("expired")

        if api_token != APP_TOKEN:
            return False, "❌ Data tidak sesuai"

        # Parse expired time (format dari API: 2026-09-01T17:00:00.000Z)
        expired = datetime.fromisoformat(expired_str.replace("Z", "+00:00"))

        # Ambil waktu real dari NTP
        current = get_internet_time()

        if current > expired:
            return False, f"❌ Automation sudah tidak valid"

        return True, "✅ License OK"

    except Exception as e:
        return False, f"❌ Gagal cek license: {e}"