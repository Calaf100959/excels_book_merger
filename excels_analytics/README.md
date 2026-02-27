# Excel Merger (Python / tkinter / Excel COM)

フォルダ内の Excel ファイル（`.xls`, `.xlsx`, `.xlsm`, `.xlsb`）を列挙し、**全シートを 1 つの Excel ファイルに統合（シートコピー方式）**する Windows 専用ツールです。
保存は統合完了後に「名前を付けて保存」を促します。

## 前提

- Windows
- Microsoft Excel（デスクトップ版）がインストールされていること

## セットアップ

基本は標準ライブラリ（`tkinter`）＋ PowerShell の Excel COM を使うため、追加インストール不要です。

## 起動

```powershell
py .\excel_merger_gui.py
```

## （任意）Python から直接 COM 操作したい場合

環境により `pip` が使えない場合があります。その場合は PowerShell 経由のまま利用してください。

```powershell
py -m pip install -r requirements.txt
```

## 仕様（現状）

- 取り込み対象: 指定フォルダ直下の Excel ファイル（`~$` で始まる一時ファイルは除外）
- 統合方式: 各ブックの **全ワークシートをコピー**して 1 ブックにまとめる
- 同名シート: 連番付与（例: `Sheet1`, `Sheet1_2`, `Sheet1_3`…）
- 保存: 統合完了後に保存ダイアログ（キャンセル時は再選択 or 保存せず終了）
