param (
    [string]$path= "",
    [string]$output= "output.docx",
    [string]$logfile= "logfile.txt"
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

python main.py --path $path --output $output --logfile $logfile

deactivate