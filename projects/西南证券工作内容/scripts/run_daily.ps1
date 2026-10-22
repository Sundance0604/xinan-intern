param([string]$CondaEnvironment = 'pavane', [string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $projectRoot 'local_data' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$day = Get-Date -Format 'yyyy-MM-dd'
$json = Join-Path $OutputDirectory 'snapshot.json'
$holdings = Join-Path $OutputDirectory ("持仓表-{0}.xlsx" -f $day)
& node (Join-Path $PSScriptRoot 'extract_qq_docs.js') --config (Join-Path $projectRoot 'config.local.json') --out $json
if ($LASTEXITCODE -ne 0) { throw 'Tencent document export failed' }
& conda run --no-capture-output -n $CondaEnvironment python (Join-Path $PSScriptRoot 'json_to_xlsx.py') $json $holdings
if ($LASTEXITCODE -ne 0) { throw 'Workbook conversion failed' }
& conda run --no-capture-output -n $CondaEnvironment python (Join-Path $PSScriptRoot 'holding_days_reminder.py') $holdings --date $day --out (Join-Path $OutputDirectory ("持仓工作日提醒_{0}.txt" -f $day))
if ($LASTEXITCODE -ne 0) { throw 'Holding reminder generation failed' }
