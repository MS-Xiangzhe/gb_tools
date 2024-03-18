param (
    [string]$path= "",
    [string]$output= "output.docx",
    [string]$logfile= "logfile.txt",
    [string[]]$extra= @(),
    [switch]$extraOnly= $false,
    [switch]$y= $false
)

if (!(Get-Command -Name python)) {
    Write-Host "Python not installed, installing..."
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

$extraArgs = ""
if ($extra) {
    foreach ($arg in $extra) {
        $extraArgs += " --extra $arg"
    }
}

$command = "python main.py --path '$path' --output '$output' --logfile '$logfile'"
if ($extraOnly) {
    $command += " --extra-only"
}
if ($y) {
    $command += " -y"
}
$command += $extraArgs

Write-Host $command
Invoke-Expression $command

deactivate