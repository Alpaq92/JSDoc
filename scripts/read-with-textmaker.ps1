<#
  read-with-textmaker.ps1 — read a .doc by driving SoftMaker TextMaker via COM.

  This is the oracle used to confirm that textToDoc()'s output opens cleanly in
  a real word processor (not just our lenient parsers). TextMaker exposes a
  Word-like automation model; we open the file and read every paragraph's text
  back. A broken .doc would fail to open, fall into recovery, or return garbage.

  Windows + SoftMaker FreeOffice/Office (TextMaker) only. Usage:
    powershell -File scripts/read-with-textmaker.ps1 -Path samples\some.doc
#>
param([Parameter(Mandatory = $true)][string]$Path)

$tm = $null
try {
  $tm = New-Object -ComObject TextMaker.Application
  try { $tm.DisplayAlerts = $false } catch {}
  $tm.Visible = $false
  $doc = $tm.Documents.Open((Resolve-Path $Path).Path)
  if ($null -eq $doc) { throw "Documents.Open returned null (TextMaker could not open it)" }
  $sb = New-Object System.Text.StringBuilder
  $n = $doc.Paragraphs.Count
  for ($i = 1; $i -le $n; $i++) {
    [void]$sb.AppendLine($doc.Paragraphs.Item($i).Range.Text)
  }
  $doc.Close($false)
  Write-Output ("paragraphs: {0}" -f $n)
  Write-Output "----"
  Write-Output $sb.ToString()
} catch {
  Write-Output ("FAILED: " + $_.Exception.Message)
  exit 1
} finally {
  if ($tm) { try { $tm.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($tm) }
  Get-Process TextMaker -ErrorAction SilentlyContinue | Stop-Process -Force
}
