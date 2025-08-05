param (
    [string]$FolderPath = ".",
    [string]$FileTypes = "*"
)

# Normalize file type filters
$fileTypePatterns = @()
if ($FileTypes -ne "*") {
    $FileTypes.Split(",") | ForEach-Object {
        $ext = $_.Trim()
        if (-not $ext.StartsWith(".")) { $ext = ".$ext" }
        $fileTypePatterns += "*$ext"
    }
} else {
    $fileTypePatterns = @("*")
}

# Create log and CSV file paths
$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "MCMIPAuditLog$timestamp.txt"
$csvFile = Join-Path -Path $PSScriptRoot -ChildPath "MCMIPAuditReport$timestamp.csv"
$failuresCsvFile = Join-Path -Path $PSScriptRoot -ChildPath "MCMIPRemovalFailures$timestamp.csv"

# Initialize CSV data arrays
$auditResults = @()
$removalFailures = @()

function Log {
    param ([string]$message)
    $entry = "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $message"
    Write-Host $entry
    Add-Content -Path $logFile -Value $entry
}

# Check if the folder exists
try {
    $f = Get-Item -Path $FolderPath -ErrorAction Stop
    if (-not $f.PSIsContainer) {
        Log """$FolderPath"" is not a folder."
        exit
    }
    $FolderPath = $f
} catch {
    Log "The folder path $FolderPath does not exist."
    exit
}

# Import the PurviewInformationProtection module
try {
    Import-Module PurviewInformationProtection -ErrorAction Stop
    Log "PurviewInformationProtection module loaded successfully."
} catch {
    Log "PurviewInformationProtection module is not installed or failed to load."
    exit
}

# Global variable to cache authentication status
if (-not $Global:MIPAuthInitialized) {
    try {
        Set-Authentication
        $Global:MIPAuthInitialized = $true
        Log "Authentication successful and cached for this session."
    } catch {
        Log "Authentication failed. Please ensure you have access to Microsoft Purview Information Protection."
        exit
    }
} else {
    Log "Reusing cached authentication for this session."
}

# Process all files recursively in a single pass
Log "Scanning folder tree: $FolderPath with filters: $($fileTypePatterns -join ', ')"
Get-ChildItem -Path $FolderPath -File -Recurse | Where-Object {
    $file = $_
    $fileTypePatterns | ForEach-Object { if ($file.Name -like $_) { return $true } }
    return $false
} | ForEach-Object {
    $filePath = $_.FullName
    $fileName = $_.Name
    $label = ""
    $labelRemoved = $false
    $errorMessage = ""

    try {
        $status = Get-FileStatus -Path $filePath
        $label = $status.MainLabelName
        $isLabeled = $status.IsLabeled

        if ($isLabeled) {
            $l = ", Labelled"
        } else {
            $l = ", Unlabelled"
        }

        Log "File: $filePath$l"

        if ($canRemoveLabel -and -not [string]::IsNullOrEmpty($label)) {
            try {
                Set-FileLabel -Path .\Book1.xlsx -LabelId 87ba5c36-b7cf-4793-bbc2-bd5b3a9f95ca -JustificationMessage "Previous label was incorrect"
                $labelRemoved = $true
                Log "  → Label removed successfully."
            } catch {
                $errorMessage += "Failed to remove label. "
                Log "  → Failed to remove label: $_"
            }
        }
    } catch {
        $errorMessage = "Error retrieving or modifying label/protection info: $_"
        Log "  → $errorMessage"
    }

    # Add to audit results
    $auditResults += [PSCustomObject]@{
        FilePath            = $filePath
        Label               = $label
        LabelRemoved        = $labelRemoved
        Error               = $errorMessage
    }

    # Add to failure report if needed
    if (($label -or $isProtected) -and $errorMessage) {
        $removalFailures += [PSCustomObject]@{
            FilePath  = $filePath
            Label     = $label
            Reason    = $errorMessage.Trim()
        }
    }
}

# Export CSVs
$auditResults | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
$removalFailures | Export-Csv -Path $failuresCsvFile -NoTypeInformation -Encoding UTF8

Log "Scan complete."
Log "Log saved to: $logFile"
Log "CSV report saved to: $csvFile"
Log "Failures report saved to: $failuresCsvFile"
