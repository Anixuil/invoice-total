param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "dist")
)

$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path $PSScriptRoot).Path
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$packageName = "invoice-total-docker-deploy-$timestamp"
$stagingDirectory = Join-Path ([System.IO.Path]::GetTempPath()) $packageName
$resolvedOutputDirectory = (New-Item -ItemType Directory -Force -Path $OutputDirectory).FullName
$archivePath = Join-Path $resolvedOutputDirectory "$packageName.zip"

$requiredFiles = @(
    "Dockerfile",
    "docker-compose.yml",
    ".dockerignore",
    "requirements.txt",
    "server.py",
    "invoice_total.py",
    "jira_processor.py",
    "weekly_report_processor.py"
)

$requiredDirectories = @(
    "static",
    "templates"
)

try {
    foreach ($relativePath in $requiredFiles) {
        $sourcePath = Join-Path $projectRoot $relativePath
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw "缺少 Docker 部署文件: $relativePath"
        }
    }

    foreach ($relativePath in $requiredDirectories) {
        $sourcePath = Join-Path $projectRoot $relativePath
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Container)) {
            throw "缺少 Docker 部署目录: $relativePath"
        }
    }

    New-Item -ItemType Directory -Force -Path $stagingDirectory | Out-Null

    foreach ($relativePath in $requiredFiles) {
        Copy-Item -LiteralPath (Join-Path $projectRoot $relativePath) -Destination (Join-Path $stagingDirectory $relativePath)
    }

    foreach ($relativePath in $requiredDirectories) {
        Copy-Item -LiteralPath (Join-Path $projectRoot $relativePath) -Destination (Join-Path $stagingDirectory $relativePath) -Recurse
    }

    if (Test-Path -LiteralPath $archivePath) {
        Remove-Item -LiteralPath $archivePath -Force
    }
    Compress-Archive -Path (Join-Path $stagingDirectory "*") -DestinationPath $archivePath -CompressionLevel Optimal
    Write-Host "Docker 部署包已生成: $archivePath"
}
finally {
    if (Test-Path -LiteralPath $stagingDirectory) {
        Remove-Item -LiteralPath $stagingDirectory -Recurse -Force
    }
}
