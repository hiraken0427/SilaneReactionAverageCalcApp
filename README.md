# SilaneReactionAverageCalcApp by hiraken0427
# 松永研究室 ゾルゲル班用解析ツール (Matsunaga Lab - Sol-Gel Team Analysis Tool)

このアプリケーションは、XLSXまたはCSV形式の時系列データ（特に電流データ）を読み込み、グラフを生成・保存するためのGUIツールです。

This is a GUI application designed to read time-series data (specifically current data) from XLSX or CSV files, generate graphs, and save them.

---

### 概要

XLSXまたはCSVファイル（Elapsed Time (s), Current (A) の2列データ）を読み込み、グラフをPNG画像として出力します。
また、指定した閾値（時間および電流）を超えるデータの平均値を計算し、グラフ化する機能を提供します。

### 主な機能

* **グラフの表示と保存**: 読み込んだデータ（全範囲）をグラフ化し、タブで表示します。グラフはPNGとしてデスクトップの `shilane_result` フォルダに自動保存されます。
* **ピーク値の解析**: 「メニュー > ピーク値平均計算」から、以下の詳細な解析が可能です。
    * **縦軸（電流）の閾値指定**: 指定した電流値（例: `2.5E-06` A）より大きいデータを抽出します。
    * **横軸（時間）の閾値指定**: 指定した時間（例: `12.0` s）以降のデータを抽出します。
    * **平均値の計算**: 上記2つの条件を満たすデータの平均値を計算し、メイングラフ上にテキストで表示します。
    * **ピークグラフの保存**: 抽出されたデータのみのグラフを、`_peak.png` という接尾辞で別途保存します。
* **設定の記憶**: 最後に開いたファイルのパスを `Documents/silane_ave_app.config` に保存し、次回起動時にその場所をデフォルトで開きます。

### 要件（必要なライブラリ）

* Python 3.x
* PySide6
* pandas
* matplotlib
* openpyxl

------------------------------------------------------------------------

# SilaneReactionAverageCalcApp by hiraken0427
# Matsunaga Lab - Sol-Gel Team Analysis Tool

This application is a GUI tool designed to read time-series data (specifically current data) from XLSX or CSV files, generate graphs, and save them.

This is a GUI application designed to read time-series data (specifically current data) from XLSX or CSV files, generate graphs, and save them.

---

### Overview

It reads XLSX or CSV files (containing two columns: Elapsed Time (s) and Current (A)) and outputs graphs as PNG images.
It also calculates and graphs the average values of data exceeding specified thresholds (both time and current).

### Main Features

* **Graph Display and Saving**: Graphs the loaded data (entire range) and displays it in tabs. Graphs are automatically saved as PNG files in the `shilane_result` folder on the desktop.
* **Peak Value Analysis**: Access detailed analysis via “Menu > Peak Value Average Calculation”:
    * **Vertical Axis (Current) Threshold Setting**: Extracts data exceeding a specified current value (e.g., `2.5E-06` A).
    * **Horizontal Axis (Time) Threshold Setting**: Extracts data from the specified time point onward (e.g., `12.0` s).
    * **Average Value Calculation**: Calculates the average value of data meeting both conditions and displays it as text on the main graph.
    * **Peak Graph Saving**: Saves a graph of only the extracted data separately with the suffix `_peak.png`.
* **Saving Settings**: Saves the path of the last opened file to `Documents/silane_ave_app.config` and opens that location by default on next launch.

### Requirements (Required Libraries)

* Python 3.x
* PySide6
* pandas
* matplotlib
* openpyxl

Translated with DeepL.com (free version)