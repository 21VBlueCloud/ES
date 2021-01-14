# script to process before import module.

# We need to check the password storing path in $home is exist or not.
# Create the path if it is not exist.

$Path = Join-Path $home "MyPassword"

if (!(Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
}

# create Json file.
$FilePath = Join-Path $Path $env:COMPUTERNAME

$Global:SecretPath = $FilePath
