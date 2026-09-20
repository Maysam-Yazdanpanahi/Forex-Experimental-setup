# Forex-Experimental-setup: Audio-Enhanced GUI for Forex Algorithmic Trading

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.x](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)

## Overview
This repository contains the **Audio-Enhanced Graphical User Interface (GUI)** developed as a comprehensive experimental setup for a Ph.D. dissertation titled: **"An Algorithmic Trading Meta-Model Based on Weighted Ensemble Technical Strategies"**.

The framework, known as **BEP-CAVE** (Binary Encoded Pattern with Confidence-Adapted Volume Enhancement), is designed to evaluate algorithmic trading strategies in the Forex market. This GUI facilitates the entire research lifecycle: from data acquisition (MetaTrader 5 / Yahoo Finance) and feature engineering to model training, backtesting, and result visualization.

## Key Features
Based on the dissertation methodology (Chapter 3), this GUI provides:
- **Data Management:** Connect live to MetaTrader 5 (MT5) or fetch historical data via Yahoo Finance. Supports custom date ranges, timeframes, and offline CSV/Excel loading.
- **Feature Engineering:** Extensive support for technical indicators (RSI, MACD, SMA, EMA, Bollinger Bands, Ichimoku, ADX, etc.) and binary pattern encoding (BEP-CAVE framework).
- **Model Configuration:** Select and run up to **23 Machine Learning and Deep Learning models** simultaneously, including Random Forest, SVM, LSTM, XGBoost, LightGBM, CatBoost, and the novel GSE (Greedy Storm Ensemble) strategy.
- **Dynamic Risk Management:** Implements Dynamic Lottage (Position Sizing) and Adaptive Money Management based on Drawdown and Confidence-Weighted Returns (CWR).
- **Validation Protocol:** Built-in support for the **DSTS (Dynamic Single Test Splitting)** protocol to prevent look-ahead bias and ensure robust out-of-sample testing.
- **Audio-Enhanced Interaction:** Features text-to-speech (TTS) and audio alarms for real-time trade signals (Buy/Sell/Hold), Take Profit, and Stop Loss alerts. Supports voice commands.
- **Comprehensive Visualization:** Pre-run and post-run visualizations, including equity curves, drawdown plots, confusion matrices, ROC curves, trade histograms, and feature correlation heatmaps.
- **Trading Simulator:** Paper trading (demo account) support and automated order execution via MetaTrader 5.
- **Preset Management:** Save and load experimental setups (features, models, hyperparameters) for reproducibility.

## Project Structure
- `PHD_Thesis_Audio_Enhanced_GUI_Ver2.py`: Main executable script containing the complete GUI and logic.
- `Custom-libs/`: Contains custom libraries (`ALARM.py` for audio alerts, `customized.py` for colored terminal outputs).
- `pictures/`: Images and icons used in the GUI interface.
- `dataset/` & `Data/`: Default directories for saving and loading datasets and preset configurations.

## Requirements & Installation
To run this GUI, ensure you have **Python 3.x** installed. You can install the required libraries using pip:

```bash
pip install pandas numpy tkinter MetaTrader5 yfinance matplotlib seaborn pyttsx3 sounddevice wavio speech_recognition pytz winsound scikit-learn tensorflow keras xgboost lightgbm catboost openpyxl screeninfo PyQt5
```

*(Note: `tkinter` is usually included by default with Python installations. `MetaTrader5` is only required if you intend to connect to the MT5 terminal.)*

## Usage
1.  Clone the repository to your local machine.
2.  Open a terminal and navigate to the project folder.
3.  Run the main script:
    ```bash
    python PHD_Thesis_Audio_Enhanced_GUI_Ver2.py
    ```
4.  Use the GUI menus to:
    - Load your database (MT5 or Yahoo Finance).
    - Select your desired features (e.g., Binary Indicators, OHLC).
    - Choose your label (Market, Rise/Fall, Scaled Return).
    - Select one or more base models (e.g., RF, LSTM).
    - Run the experiment and view the generated reports.

## Citation
If you use this code or the BEP-CAVE framework in your research, please cite the corresponding Ph.D. dissertation:
> **Yazdanpanahi, M.** (2026). *An Algorithmic Trading Meta-Model Based on Weighted Ensemble Technical Strategies* (Ph.D. Thesis). Shahrood University of Technology, Faculty of Computer Engineering.

## License & Intellectual Property
This project is licensed under the **MIT License**. 
*Note: According to the author's university regulations, the intellectual property rights of this work belong to Shahrood University of Technology. The code is shared for academic and research purposes.*

## Acknowledgements
- Supervisor: Dr. Morteza Zahedi
- Advisor: Dr. Mohammad Mahdi Hosseini
