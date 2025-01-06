param (
    [string]$path = "",
    [string]$output = "output.docx",
    [string]$logfile = "logfile.txt",
    [switch]$docRange = $false,
    [string]$preExtra = "",
    [string]$extra = "",
    [switch]$extraOnly = $false,
    [switch]$passText = $false,
    [switch]$y = $false,
    [switch]$skipChange = $false,
    [switch]$askGuessReplace = $false
)
$pythonVersion = & python --version 2>&1
$pythonVersion = $pythonVersion -replace 'Python ', ''

if ($pythonVersion -notmatch "^3\.11|^3\.12|^3\.13") {
    Write-Host "Python version does not start with 3.11, 3.12 or 3.13, installing..."
    winget.exe install python
    Write-Host "Python installed, you need restart your terminal"
    exit
}

# if not have .venv folder
if (!(Test-Path .venv)) {
    Write-Host "Creating virtual environment..."
    python -m venv .venv
}

.\.venv\Scripts\Activate
pip install -r requirements.txt

$command = "python main.py --path '$path' --output '$output' --logfile '$logfile'"
if ($docRange) {
    $command += " --doc-range"
}
if ($extraOnly) {
    $command += " --extra-only"
}
if ($passText) {
    $command += " --pass-text"
}
if ($y) {
    $command += " -y"
}

if ($skipChange) {
    $command += " --skip-change"
}

if ($askGuessReplace) {
    $command += " --ask-guess-replace"
}

$extraArgs = ""
if ($preExtra) {
    $extraArray = $preExtra -split ","
    $extraArgs += " --pre-extra"
    foreach ($arg in $extraArray) {
        $extraArgs += " $arg"
    }
}
$command += $extraArgs

$extraArgs = ""
if ($extra) {
    $extraArray = $extra -split ","
    $extraArgs += " --extra"
    foreach ($arg in $extraArray) {
        $extraArgs += " $arg"
    }
}
$command += $extraArgs

Write-Host $command
Invoke-Expression $command

deactivate