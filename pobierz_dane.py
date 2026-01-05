import xlwings as xw
import yfinance as yf
import pandas as pd
from datetime import date

def main():
    # Pobieramy aktywny skoroszyt (to zadziała zarówno z Excela, jak i z VS Code)
    try:
        wb = xw.Book.caller()
    except Exception:
        # Jeśli uruchamiasz ręcznie z VS Code, połączy się z otwartym plikiem
        wb = xw.books.active 

    ws = wb.sheets['DATA']
    
    # 1. Pobieramy Ticker z H1 (tak jak chciałaś)
    ticker = ws.range('H1').value 
    
    if not ticker:
        ticker = "TSLA"

    # 2. Pobieranie danych (Twój oryginał)
    print(f"Pobieram dane dla: {ticker}...")
    dane = yf.download(ticker, start="2000-01-01", end=date.today(), auto_adjust=True)
    
    if dane.empty:
        print("Błąd: Nie pobrano danych!")
        return

    # 3. Naprawa MultiIndex (Twój oryginał)
    if isinstance(dane.columns, pd.MultiIndex):
        dane.columns = dane.columns.get_level_values(0)

    dane.reset_index(inplace=True)
    dane['Date'] = dane['Date'].dt.date
    dane = dane[['Date', 'Open', 'High', 'Low', 'Close', 'Volume']]

    # 4. WKLEJANIE (Poprawione nazwy zmiennych i zakres)
    # Zmieniamy ws.clear() na ws.range('A:F').clear_contents(), 
    # żeby nie skasować ustawień w H1-H4!
    ws.range('A:F').clear_contents()
    
    # Wklejamy dane (używamy ws, bo tak zdefiniowaliśmy na początku)
    ws.range('A1').options(index=False).value = dane
    
    print(f"Sukces! Pbrano {len(dane)} wierszy.")

if __name__ == "__main__":
    main()