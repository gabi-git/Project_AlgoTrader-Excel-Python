import sys
import os
import requests
from bs4 import BeautifulSoup
from textblob import TextBlob

# --- 1. KONFIGURACJA ŚCIEŻKI ---
# Plik wynikowy zapisze się w tym samym folderze, gdzie jest ten skrypt
FOLDER_SKRYPTU = os.path.dirname(os.path.abspath(__file__))
output_file = os.path.join(FOLDER_SKRYPTU, "market_mood.txt")

# --- 2. ODBIERAMY TICKER OD EXCELA ---
# Excel wyśle nam np. "AAPL" jako argument.
if len(sys.argv) > 1:
    ticker_symbol = sys.argv[1]
else:
    ticker_symbol = "NVDA" # Domyślnie, jeśli uruchomisz ręcznie bez niczego

print(f"--- ANALIZA SENTYMENTU DLA: {ticker_symbol} ---")

def get_google_sentiment(ticker):
    # Szukamy newsów
    url = f"https://news.google.com/rss/search?q={ticker}+stock+news&hl=en-US&gl=US&ceid=US:en"
    try:
        response = requests.get(url, timeout=5)
        soup = BeautifulSoup(response.content, features="xml")
        items = soup.findAll('item')
        
        if not items:
            soup = BeautifulSoup(response.content, features="html.parser")
            items = soup.findAll('item')
        
        if not items: return 0, "No news found", "Empty"

        # Słowa kluczowe
        bullish_words = ["surge", "jump", "soar", "record", "high", "beat", "strong", "buy", "gain", "profit", "bull", "rally"]
        bearish_words = ["plunge", "crash", "drop", "miss", "loss", "weak", "sell", "bear", "down", "low", "fall", "tumble"]

        total_score = 0
        count = 0
        latest_headline = ""

        # Analiza pierwszych 10 nagłówków
        for index, item in enumerate(items):
            if index > 9: break
            title = item.title.text
            if index == 0: latest_headline = title
            
            blob_score = TextBlob(title).sentiment.polarity
            boost = 0
            lower_title = title.lower()
            
            for word in bullish_words:
                if word in lower_title: boost += 0.3
            for word in bearish_words:
                if word in lower_title: boost -= 0.3
            
            total_score += (blob_score + boost)
            count += 1

        if count > 0:
            avg = max(-1, min(1, total_score / count))
            return avg, latest_headline, "OK"
        return 0, "No News Data", "Empty"

    except Exception as e:
        return 0, "Error: " + str(e), "Error"

# --- 3. URUCHOMIENIE I ZAPIS ---
score, headline, status = get_google_sentiment(ticker_symbol)

# Ustalanie słownego nastroju
if score >= 0.1: mood = "Positive (Bullish)"
elif score <= -0.1: mood = "Negative (Bearish)"
else: mood = "Neutral"

# Formatowanie wyniku (kropka w liczbie)
score_display = f"{score:.2f}"

# Zapis do pliku
try:
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(f"{mood}\n")
        f.write(f"{score_display}\n")
        # Skracamy nagłówek, żeby nie rozwalił raportu
        clean_headline = headline.replace('\n', ' ').strip()[:80] + "..."
        f.write(f"{clean_headline}\n")
    print(f"SUKCES: Zapisano dla {ticker_symbol}")
except Exception as e:
    print(f"BŁĄD ZAPISU: {e}")