param(
  [Parameter(Mandatory=$true)]
  [string]$FileListPath,

  [Parameter(Mandatory=$true)]
  [string]$CancelFlagPath,

  [Parameter(Mandatory=$false)]
  [string]$SuggestedName = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Ensure UTF-8 for stdin/stdout when called from Python.
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
[Console]::InputEncoding = $utf8
$OutputEncoding = $utf8

function Write-Log([string]$Message) {
  Write-Output ("LOG|" + $Message)
}

function Write-ProgressLine([int]$Current, [int]$Total, [string]$Filename) {
  Write-Output ("PROGRESS|{0}|{1}|{2}" -f $Current, $Total, $Filename)
}

function Sanitize-SheetName([string]$Name) {
  $n = $Name.Trim()
  if ([string]::IsNullOrWhiteSpace($n)) { $n = "Sheet" }
  $illegal = @(":", "\", "/", "?", "*", "[", "]")
  foreach ($ch in $illegal) { $n = $n.Replace($ch, "_") }
  if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
  return $n
}

function Get-UniqueSheetName([string[]]$ExistingNames, [string]$Desired) {
  $base = Sanitize-SheetName $Desired
  if (-not ($ExistingNames -contains $base)) { return $base }

  for ($i = 2; $i -lt 10000; $i++) {
    $suffix = "_" + $i
    $maxBase = 31 - $suffix.Length
    $candBase = $base
    if ($candBase.Length -gt $maxBase) { $candBase = $candBase.Substring(0, $maxBase) }
    $cand = $candBase + $suffix
    if (-not ($ExistingNames -contains $cand)) { return $cand }
  }
  throw "Failed to generate a unique sheet name."
}

function Get-FileFormat([string]$Path) {
  $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
  switch ($ext) {
    ".xlsx" { return 51 } # xlOpenXMLWorkbook
    ".xlsm" { return 52 } # xlOpenXMLWorkbookMacroEnabled
    ".xlsb" { return 50 } # xlExcel12
    ".xls"  { return 56 } # xlExcel8
    default { return 51 }
  }
}

$excel = $null
$destWb = $null

try {
  $files = Get-Content -LiteralPath $FileListPath -Encoding UTF8 | Where-Object { $_ -and $_.Trim().Length -gt 0 }
  $files = @($files)
  if ($files.Count -eq 0) {
    Write-Log "対象ファイルがありません。"
    exit 1
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.EnableEvents = $false
  try { $excel.AutomationSecurity = 3 } catch {}
  try { $excel.Calculation = -4135 } catch {} # xlCalculationManual

  $destWb = $excel.Workbooks.Add()
  $initialSheetNames = @()
  for ($i = 1; $i -le $destWb.Worksheets.Count; $i++) {
    $initialSheetNames += [string]$destWb.Worksheets.Item($i).Name
  }

  $totalFiles = $files.Count
  Write-Log ("対象ファイル数: {0}" -f $totalFiles)

  $copiedAny = $false
  $missing = [System.Reflection.Missing]::Value

  for ($fi = 0; $fi -lt $totalFiles; $fi++) {
    if (Test-Path -LiteralPath $CancelFlagPath) {
      Write-Log "キャンセルされました。後処理中..."
      exit 2
    }

    $path = $files[$fi]
    $name = [System.IO.Path]::GetFileName($path)
    Write-ProgressLine ($fi + 1) $totalFiles $name
    Write-Log ("開く: {0}" -f $name)

    $srcWb = $null
    try {
      $srcWb = $excel.Workbooks.Open($path, 0, $true)
      $wsCount = $srcWb.Worksheets.Count

      for ($wi = 1; $wi -le $wsCount; $wi++) {
        if (Test-Path -LiteralPath $CancelFlagPath) {
          Write-Log "キャンセルされました。後処理中..."
          exit 2
        }

        $srcWs = $srcWb.Worksheets.Item($wi)
        $desired = [string]$srcWs.Name

        $after = $destWb.Worksheets.Item($destWb.Worksheets.Count)
        $null = $srcWs.Copy($missing, $after)
        $copied = $excel.ActiveSheet

        $existing = @()
        for ($j = 1; $j -le $destWb.Worksheets.Count; $j++) {
          $existing += [string]$destWb.Worksheets.Item($j).Name
        }
        # Exclude the just-copied sheet's current name from collision set.
        $existing = $existing | Where-Object { $_ -ne [string]$copied.Name }

        $unique = Get-UniqueSheetName $existing $desired
        try {
          $copied.Name = $unique
        } catch {
          $fallback = Get-UniqueSheetName $existing ($desired + "_copy")
          $copied.Name = $fallback
        }

        $copiedAny = $true
        Write-Log ("  コピー: {0} -> {1}" -f $desired, [string]$copied.Name)
      }
    } catch {
      Write-Log ("[WARN] {0} を処理できません: {1}" -f $name, $_.Exception.Message)
    } finally {
      try { if ($srcWb -ne $null) { $srcWb.Close($false) } } catch {}
      try { if ($srcWb -ne $null) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($srcWb) } } catch {}
      $srcWb = $null
    }
  }

  if ($copiedAny -and $initialSheetNames.Count -gt 0) {
    foreach ($shName in $initialSheetNames) {
      try {
        if ($destWb.Worksheets.Count -le 1) { break }
        $destWb.Worksheets.Item($shName).Delete()
      } catch {}
    }
  }

  Write-Log "統合完了。保存先を選択してください。"
  $suggested = $SuggestedName
  if ([string]::IsNullOrWhiteSpace($suggested)) {
    $suggested = ("merged_{0}_{1}.xlsx" -f $destWb.Name, (Get-Date -Format "yyyyMMdd_HHmmss"))
  }
  Write-Output ("REQUEST_SAVE|" + $suggested)

  $savePath = [Console]::In.ReadLine()
  if ([string]::IsNullOrWhiteSpace($savePath)) {
    Write-Log "保存がキャンセルされました。保存せずに終了します。"
    exit 3
  }

  Write-Log ("SavePath(Received): {0}" -f $savePath)
  $savePath = $savePath.Trim()
  $savePath = $savePath -replace '/', '\'
  if ([string]::IsNullOrWhiteSpace([System.IO.Path]::GetExtension($savePath))) {
    $savePath = $savePath + ".xlsx"
  }
  try {
    $savePath = [System.IO.Path]::GetFullPath($savePath)
  } catch {}
  Write-Log ("SavePath(Normalized): {0}" -f $savePath)

  $fmt = Get-FileFormat $savePath
  $destWb.SaveAs($savePath, $fmt)
  Write-Log ("保存しました: {0}" -f $savePath)
  exit 0
}
catch {
  try {
    $msg = $_.Exception.Message
    if ([string]::IsNullOrWhiteSpace($msg)) { $msg = [string]$_ }
    Write-Log ("[ERROR] " + $msg)
  } catch {}
  exit 1
}
finally {
  try { if ($destWb -ne $null) { $destWb.Close($false) } } catch {}
  try { if ($excel -ne $null) { $excel.Quit() } } catch {}
  try { if ($destWb -ne $null) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($destWb) } } catch {}
  try { if ($excel -ne $null) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } } catch {}
  $destWb = $null
  $excel = $null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

