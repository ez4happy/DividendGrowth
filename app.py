import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
import time

# ===== 최신월 자동 POSITION 계산 함수 =====
def get_latest_band_prices_robust(gicode, max_retries=2, delay=0.2):
    url = (
        f"https://comp.fnguide.com/SVO2/common/chartListPopup2.asp"
        f"?oid=pbrBandCht&cid=01_06&gicode={gicode}&filter=D&term=Y&etc=B&etc2=0"
    )
    latest_dt = None
    latest_cells = None

    for _ in range(max_retries):
        try:
            resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
            resp.encoding = 'utf-8'
            if resp.status_code != 200:
                time.sleep(delay)
                continue
            soup = BeautifulSoup(resp.text, 'html.parser')
            rows = soup.select("table tbody tr")
            for tr in rows:
                cells = [td.get_text(strip=True).replace(',', '') for td in tr.find_all('td')]
                if len(cells) >= 7:
                    dt = cells[0]
                    if len(dt) == 10 and dt.replace('/', '').isdigit():
                        if latest_dt is None or dt > latest_dt:
                            latest_dt = dt
                            latest_cells = cells
            if latest_cells:
                try:
                    adj_price = float(latest_cells[1]) if latest_cells[1] != '-' else None
                except:
                    adj_price = None
                band_prices = []
                for val in latest_cells[2:7]:
                    if val != '-':
                        band_prices.append(float(val))
                if len(band_prices) == 5:
                    return latest_dt, adj_price, band_prices
            time.sleep(delay)
        except:
            time.sleep(delay)
    return None, None, None

def get_position_from_price(price, band_prices):
    if price is None or np.isnan(price) or band_prices is None or len(band_prices) != 5:
        return np.nan
    if price < band_prices[0]:
        return 1
    for i in range(1, 5):
        if band_prices[i-1] <= price < band_prices[i]:
            return i + 1
    return 6

# ===== 엑셀 로드 =====
df = pd.read_excel("1.xlsx")

# ===== POSITION 계산 =====
df['position'] = np.nan
for idx, row in df.iterrows():
    gicode = f"A{str(row['종목코드']).zfill(6)}"
    latest_date, _, bands = get_latest_band_prices_robust(gicode)
    pos = get_position_from_price(row['현재가'], bands)
    df.at[idx, 'position'] = pos
    print(f"[{row['종목명']}] 최신월: {latest_date}, bands={bands}, 현재가={row['현재가']}, position={pos}")
    time.sleep(0.1)  # 서버 예의

# ===== 저장 =====
df.to_excel("result.xlsx", index=False)
print("✅ position 계산 완료 → result.xlsx 저장됨")
