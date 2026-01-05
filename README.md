
# AlgoTrader - Hybrid Investment Analysis System

## Project Overview
AlgoTrader is a hybrid financial analysis tool that integrates **Microsoft Excel (VBA)** with **Python**. The system automates the process of retrieving historical stock data, performing technical analysis (SMA, RSI), and conducting AI-based sentiment analysis using Natural Language Processing (NLP) on real-time news.

The tool simulates investment strategies against a Buy & Hold benchmark and generates comprehensive PDF reports with risk metrics (Sharpe Ratio, Volatility) and actionable recommendations ("Oracle Verdict").

## Key Features
* **Hybrid Architecture:** Excel serves as the frontend for data visualization, while Python acts as the calculation engine and data scraper.
* **Backtesting Engine:** Simulates SMA Trend Following and RSI Mean Reversion strategies.
* **AI Sentiment Analysis:** Fetches and analyzes news headlines from Google News to determine market mood (Bullish/Bearish).
* **Market Scanner (Watchlist):** Batch processing of multiple tickers to identify market opportunities and overbought/oversold conditions.
* **Reporting:** Automatic generation of professional PDF reports including technical and fundamental insights.

---

## Installation & Setup

### 1. Prerequisites
* Operating System: Windows
* Microsoft Excel (Desktop version with Macros enabled)
* Python 3.x installed (Anaconda distribution recommended)

### 2. Dependencies
Before running the tool, install the required Python libraries. Open your terminal (Command Prompt) and execute:

```bash
pip install pandas yfinance requests beautifulsoup4 textblob lxml openpyxl

```

### 3. Configuration (Critical Step)

The Excel VBA module must be linked to your local Python executable.

1. Open `AlgoTrader.xlsm`.
2. Open the VBA Editor 
3. In the Project Explorer, open **Module1**.
4. Locate the following line in the `RunRealSentimentAnalysis` and `UpdateWatchlist` subroutines:
```vba
pythonExe = "C:\Users\HP\anaconda3\python.exe"

```
5. Replace the path with the location of your Python executable.
* *Tip: To find your path, type `where python` in the Command Prompt.*


6. Save the VBA project and close the editor.

---

## Usage Guide

### Single Stock Analysis

1. Open the Excel workbook.
2. Click the **OPEN ANALYSIS** button on the main dashboard.
3. Input the stock ticker (e.g., TSLA), initial capital, and date range.
4. Click **Compare Strategies**. The system will calculate returns and risk metrics.
5. Click **Generate Report** to produce a PDF containing:
* Strategy performance comparison.
* "Oracle" investment verdict.
* AI-driven market sentiment analysis.



### Market Scanner (Watchlist)

1. Navigate to the **WATCHLIST** sheet.
2. Enter stock tickers in Column A.
3. Run the update macro. The system will populate the table with real-time price data, trend indicators, and sentiment scores for multiple assets simultaneously.

---

---

# AlgoTrader - Hybrydowy System Analizy Inwestycyjnej

## Opis Projektu

AlgoTrader to zaawansowane narzędzie analityczne łączące interfejs **Microsoft Excel (VBA)** z mocą obliczeniową języka **Python**. System automatyzuje pobieranie historycznych danych giełdowych, przeprowadza analizę techniczną (SMA, RSI) oraz analizę sentymentu rynkowego przy użyciu algorytmów przetwarzania języka naturalnego (NLP).

Narzędzie symuluje skuteczność strategii inwestycyjnych względem benchmarku "Kup i Trzymaj" oraz generuje profesjonalne raporty PDF zawierające miary ryzyka (Wskaźnik Sharpe'a, Zmienność) oraz rekomendacje inwestycyjne ("Werdykt Oracle").

## Główne Funkcjonalności

* **Architektura Hybrydowa:** Excel pełni funkcję interfejsu użytkownika, Python odpowiada za logikę obliczeniową i pobieranie danych.
* **Silnik Backtestingowy:** Symulacja strategii podążania za trendem (SMA) oraz powrotu do średniej (RSI).
* **Analiza Sentymentu AI:** Analiza nagłówków z Google News w celu określenia nastroju rynku (Pozytywny/Negatywny).
* **Skaner Rynku (Watchlist):** Seryjna analiza wielu spółek w celu identyfikacji okazji inwestycyjnych oraz stanów wykupienia/wyprzedania rynku.
* **Raportowanie:** Automatyczne generowanie raportów PDF zawierających analizę techniczną i fundamentalną.

---

## Instalacja i Konfiguracja

### 1. Wymagania systemowe

* System operacyjny: Windows
* Microsoft Excel (wersja desktopowa z włączoną obsługą makr)
* Zainstalowany Python 3.x (zalecana dystrybucja Anaconda)

### 2. Instalacja bibliotek

Przed uruchomieniem należy zainstalować wymagane biblioteki Python. W tym celu należy otworzyć terminal (CMD) i wykonać polecenie:

```bash
pip install pandas yfinance requests beautifulsoup4 textblob lxml openpyxl

```

### 3. Konfiguracja (Kluczowy krok)

Moduł VBA w Excelu musi zostać połączony z lokalną instalacją Pythona.

1. Otwórz plik `AlgoTrader.xlsm`.
2. Otwórz Edytor VBA.
3. W drzewie projektu otwórz **Module1**.
4. Znajdź linię kodu w procedurach `RunRealSentimentAnalysis` oraz `UpdateWatchlist`:
```vba
pythonExe = "C:\Users\HP\anaconda3\python.exe"

```

5. Zastąp podaną ścieżkę lokalizacją pliku wykonywalnego Python na swoim komputerze.
* *Wskazówka: Aby znaleźć ścieżkę, wpisz `where python` w wierszu poleceń.*


6. Zapisz projekt VBA i zamknij edytor.

---

## Instrukcja Użytkowania

### Analiza pojedynczej spółki

1. Otwórz skoroszyt Excel.
2. Kliknij przycisk **OPEN ANALYSIS** w głównym arkuszu.
3. Wprowadź ticker spółki (np. TSLA), kapitał początkowy oraz zakres dat.
4. Kliknij **Compare Strategies**. System obliczy stopy zwrotu i wskaźniki ryzyka.
5. Kliknij **Generate Report**, aby wygenerować plik PDF zawierający:
* Porównanie wyników strategii.
* Werdykt inwestycyjny modułu "Oracle".
* Analizę nastrojów rynkowych opartą na AI.



### Skaner Rynku (Watchlist)

1. Przejdź do arkusza **WATCHLIST**.
2. Wprowadź symbole spółek (tickery) w kolumnie A.
3. Uruchom makro aktualizujące. System automatycznie uzupełni tabelę o aktualne ceny, wskaźniki trendu oraz ocenę sentymentu dla wielu aktywów jednocześnie.

```

```