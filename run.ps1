param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $PipelineArgs
)
$ErrorActionPreference = 'Stop'
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$venvPython = Join-Path $PSScriptRoot '.venv/Scripts/python.exe'
$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
$pythonLauncher = Get-Command py -ErrorAction SilentlyContinue
if (Test-Path -LiteralPath $venvPython) {
    $pipelinePython = $venvPython
} elseif ($pythonCommand) {
    $pipelinePython = $pythonCommand.Source
} elseif ($pythonLauncher) {
    & $pythonLauncher.Source -3 (Join-Path $PSScriptRoot 'pdf_to_markdown.py') @PipelineArgs
    exit $LASTEXITCODE
} else {
    throw 'Python bulunamadi. Python 3.10+ kurun ve yerel bagimliliklari hazirlayin.'
}
& $pipelinePython (Join-Path $PSScriptRoot 'pdf_to_markdown.py') @PipelineArgs
exit $LASTEXITCODE
