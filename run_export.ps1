# run_export.ps1 — One-time script to create the alpha_roster BQ table
# Follows the exact pattern from Refresh-Data.ps1 step 11 (doom_loop)

$S3_BUCKET = "prod-academics-studient-athena-results"
$ATHENA_DB = "coachbot_feed"
$ATHENA_OUTPUT = "s3://$S3_BUCKET/athena-query-results/"
$GCP_PROJECT = "studient-flat-exports-doan"
$ROSTER_LOCAL_PATH = "C:\temp\khiem_alpha_roster\"
$ROSTER_GCS_PATH = "gs://studient-flat-exports-doan/khiem_alpha_roster"
$BQ_ROSTER_TABLE = "alpha_roster"
$PYTHON_EXE = "C:\Users\doank\AppData\Local\Programs\Python\Python312\python.exe"
$BQ_LOAD_SCRIPT = "C:\Users\doank\Documents\Projects\Studient Excel Automation\bq_load.py"

$ErrorActionPreference = "Stop"

function Run-AthenaQuery {
    param([string]$Query, [string]$Description, [int]$TimeoutSeconds = 600)
    Write-Host "Running: $Description"
    # Write query to JSON input file to preserve double-quote escaping
    $jsonInput = @{
        QueryString = $Query
        QueryExecutionContext = @{ Database = $ATHENA_DB }
        ResultConfiguration = @{ OutputLocation = $ATHENA_OUTPUT }
    } | ConvertTo-Json -Depth 3
    $jsonFile = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($jsonFile, $jsonInput, [System.Text.UTF8Encoding]::new($false))
    $queryId = aws athena start-query-execution `
        --cli-input-json "file://$jsonFile" `
        --query 'QueryExecutionId' --output text
    Remove-Item $jsonFile -ErrorAction SilentlyContinue
    if (-not $queryId) { throw "Failed to start: $Description" }
    Write-Host "Query ID: $queryId"
    $elapsed = 0
    while ($elapsed -lt $TimeoutSeconds) {
        Start-Sleep -Seconds 5; $elapsed += 5
        $result = aws athena get-query-execution --query-execution-id $queryId --output json | ConvertFrom-Json
        $status = $result.QueryExecution.Status.State
        Write-Host "  Status: $status (${elapsed}s)"
        if ($status -eq 'SUCCEEDED') { Write-Host "  Done!" -ForegroundColor Green; return }
        elseif ($status -eq 'FAILED') { throw "FAILED: $($result.QueryExecution.Status.StateChangeReason)" }
        elseif ($status -eq 'CANCELLED') { throw "CANCELLED" }
    }
    throw "Timed out after $TimeoutSeconds seconds"
}

try {
    Write-Host "`n=== ALPHA ROSTER EXPORT ===" -ForegroundColor Cyan

    # 1. Drop old table
    Run-AthenaQuery -Query "DROP TABLE IF EXISTS $ATHENA_DB.khiem_alpha_roster_export;" -Description "Drop old alpha roster table"

    # 2. Clean S3
    aws s3 rm "s3://$S3_BUCKET/exports/khiem_alpha_roster/" --recursive

    # 3. CTAS — read SQL from file to preserve double-quote escaping for reserved word "group"
    $sqlFile = Join-Path $PSScriptRoot "alpha_roster_ctas.sql"
    $ctasRoster = Get-Content -Path $sqlFile -Raw
    Run-AthenaQuery -Query $ctasRoster -Description "CTAS Alpha Roster" -TimeoutSeconds 300

    # 4. S3 → Local
    if (Test-Path $ROSTER_LOCAL_PATH) { Remove-Item -Recurse -Force $ROSTER_LOCAL_PATH }
    New-Item -ItemType Directory -Force -Path $ROSTER_LOCAL_PATH | Out-Null
    aws s3 sync "s3://$S3_BUCKET/exports/khiem_alpha_roster/" $ROSTER_LOCAL_PATH

    # 5. Local → GCS
    Write-Host "Cleaning GCS alpha roster..."
    $ErrorActionPreference = "Continue"
    gcloud storage rm --recursive "${ROSTER_GCS_PATH}/*" 2>&1 | Out-Null
    $ErrorActionPreference = "Stop"

    Write-Host "Uploading alpha roster to GCS..."
    $ErrorActionPreference = "Continue"
    gcloud storage rsync --recursive $ROSTER_LOCAL_PATH "${ROSTER_GCS_PATH}/"
    $uploadExit = $LASTEXITCODE
    $ErrorActionPreference = "Stop"
    if ($uploadExit -ne 0) { throw "GCS upload failed (exit $uploadExit)" }

    # 6. GCS → BigQuery
    Write-Host "Loading alpha_roster to BigQuery..."
    & $PYTHON_EXE -u $BQ_LOAD_SCRIPT $BQ_ROSTER_TABLE "${ROSTER_GCS_PATH}/*"
    if ($LASTEXITCODE -ne 0) { throw "BigQuery load failed" }

    Write-Host "`n=== ALPHA ROSTER EXPORT COMPLETE ===" -ForegroundColor Green
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
}
