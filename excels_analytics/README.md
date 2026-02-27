# Excel Merger (Python / tkinter / Excel COM)

フォルダ内の Excel ファイル（`.xls`, `.xlsx`, `.xlsm`, `.xlsb`）を列挙し、
**全シートを 1 つの Excel ファイルに統合（シートコピー方式）**する Windows 専用ツールです。

## 前提
- Windows
- Microsoft Excel（デスクトップ版）がインストールされていること

## 開発実行
```powershell
py .\excel_merger_gui.py
```

## EXE配布版の作成
このプロジェクトは PyInstaller で単体EXE化できます。

```powershell
cd excels_analytics
build_exe.bat
```

成功すると以下が生成されます。
- `excels_analytics\dist\ExcelMerger.exe`

## 配布時の推奨構成
配布するのは基本的に以下のみでOKです。

```text
ExcelMerger/
  ExcelMerger.exe
  README.txt (任意)
```

## 仕様（現状）
- 取り込み対象: 指定フォルダ直下の Excel ファイル（`~$` で始まる一時ファイルは除外）
- 統合方式: 各ブックの **全ワークシートをコピー**して 1 ブックにまとめる
- 同名シート: 連番付与（例: `Sheet1`, `Sheet1_2`, `Sheet1_3`…）
- 保存: 統合完了後に保存ダイアログ（キャンセル時は再選択 or 保存せず終了）

## メンテナンスメモ
- `.tmp` / `__pycache__` / `dist` / `build` はGit管理対象外
- 配布対象は `dist\ExcelMerger.exe` を基準にする
