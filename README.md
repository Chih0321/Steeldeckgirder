﻿# SteelDeck Section Calculator

計算鋼床鈑橋斷面性質

## 安裝 (install)

[SecCal.zip](T:\63427\9.鋼床鈑斷面計算\SecCal.zip)解壓縮後，執行SecCal/SecCal.exe

for programmer, clone the repository and run DSEC.py or excute the file in test folder.

## 使用教學 (tutorial)

[使用教學](https://app.heptabase.com/w/63e75860aff26baf3d7e4ec9731765ab8f41a68ccd89458ac0912b2977cc3f1c)

## 快速使用 (getting started)

全斷面計算:
使用方法分兩種1. 透過MCT指令由MIDAS計算，2. 程式直接計算
1. 使用Section.xlsx/xlsm格式填入相關尺寸細節
2. 執行SecCal.exe，路徑選擇需求輸入檔
3. 依需求執行對應按鈕
4. 由程式計算，可至Result Excel複製斷面性質
5. MCT方法，在MIDAS執行MCT中*SECTION指令，insert command，並接續貼上生成指令，執行RUN
6. MIDAS於Section table複製所得斷面性質，貼入Section.xlsx/xlsm中"斷面性質"分頁，進行shear, torison計算及格式調整

有效斷面:
程式直接計算
採道示方法，並於最後以有效斷面內插

容許應力:
程式直接計算

## 幫助 (support)

如果有任何問題，可以透過開issue 或者 可以在Message發問。

## 授權 (License)

本專案資訊請看 [LICENSE.md](LICENSE.md)

## Version Log

[CHANGELOG.md](CHANGELOG.md)
