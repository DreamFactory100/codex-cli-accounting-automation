#Requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $script:Utf8NoBom
[Console]::OutputEncoding = $script:Utf8NoBom
$OutputEncoding = $script:Utf8NoBom

$projectRoot = $PSScriptRoot
$pythonExe = Join-Path $projectRoot '.venv\Scripts\python.exe'
$runsRoot = Join-Path $projectRoot 'runs'
$runDirectory = $null

function ConvertTo-NativeArgument {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    if ($Value -notmatch '[\s"]') {
        return $Value
    }

    # Apply the Windows CommandLineToArgvW escaping rules.
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function Invoke-CapturedProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$ArgumentList,
        [AllowNull()][string]$StandardInputText = $null
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.Arguments = (($ArgumentList | ForEach-Object { ConvertTo-NativeArgument $_ }) -join ' ')
    $startInfo.WorkingDirectory = $projectRoot
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.StandardOutputEncoding = $script:Utf8NoBom
    $startInfo.StandardErrorEncoding = $script:Utf8NoBom

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) {
        throw "Could not start process: $FilePath"
    }

    # Read both streams asynchronously so a full pipe cannot deadlock the process.
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    if ($null -ne $StandardInputText) {
        $process.StandardInput.Write($StandardInputText)
    }
    $process.StandardInput.Close()
    $process.WaitForExit()

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    $exitCode = $process.ExitCode
    $process.Dispose()

    return [pscustomobject]@{
        ExitCode = $exitCode
        StdOut  = $stdout
        StdErr  = $stderr
    }
}

function Write-StageLog {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Result
    )

    $text = "EXIT CODE: $($Result.ExitCode)`r`n`r`n[STDOUT]`r`n$($Result.StdOut)" +
            "`r`n[STDERR]`r`n$($Result.StdErr)"
    [System.IO.File]::WriteAllText($Path, $text, $script:Utf8NoBom)
}

try {
    if (-not (Test-Path -LiteralPath $pythonExe -PathType Leaf)) {
        throw "Project Python was not found: .venv\Scripts\python.exe"
    }

    [System.IO.Directory]::CreateDirectory($runsRoot) | Out-Null
    do {
        $runName = (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
        $runDirectory = Join-Path $runsRoot $runName
    } while (Test-Path -LiteralPath $runDirectory)
    [System.IO.Directory]::CreateDirectory($runDirectory) | Out-Null

    $relativeRun = 'runs/' + $runName
    $workbookRelative = "$relativeRun/accounting_analysis.xlsx"
    $workbookPath = Join-Path $runDirectory 'accounting_analysis.xlsx'
    $validationPath = Join-Path $runDirectory 'validation.json'
    $reportRelative = "$relativeRun/ai_report.md"
    $reportPath = Join-Path $runDirectory 'ai_report.md'

    Write-Host '[1/4] Python accounting analysis'
    $analysis = Invoke-CapturedProcess $pythonExe @(
        '-B', 'src/analyze_transactions.py',
        '--input', 'data/abc_transactions_2026Q1.csv',
        '--output', $workbookRelative
    )
    Write-StageLog (Join-Path $runDirectory 'analysis.log') $analysis
    if ($analysis.ExitCode -ne 0) {
        throw "Accounting analysis failed (exit code $($analysis.ExitCode)); stopping."
    }
    if (-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)) {
        throw 'Analysis exited successfully but did not create the Excel workbook.'
    }

    Write-Host '[2/4] 17 unit tests'
    $tests = Invoke-CapturedProcess $pythonExe @('-B', '-m', 'unittest', 'discover', '-s', 'tests')
    Write-StageLog (Join-Path $runDirectory 'tests.log') $tests
    if ($tests.ExitCode -ne 0) {
        throw "Unit tests failed (exit code $($tests.ExitCode)); stopping."
    }

    Write-Host '[3/4] Independent validation'
    $validation = Invoke-CapturedProcess $pythonExe @(
        '-B', 'src/validate_results.py',
        '--workbook', $workbookRelative,
        '--answer-key', 'data/validation_answer_key.csv'
    )
    Write-StageLog (Join-Path $runDirectory 'validation.log') $validation
    [System.IO.File]::WriteAllText($validationPath, $validation.StdOut, $script:Utf8NoBom)
    if ($validation.ExitCode -ne 0) {
        throw "Independent validation failed (exit code $($validation.ExitCode)); stopping."
    }

    try {
        $validationObject = $validation.StdOut | ConvertFrom-Json
    }
    catch {
        throw "Validation exited successfully but produced invalid JSON: $($_.Exception.Message)"
    }
    if ($validationObject.status -ne 'PASS') {
        throw "Validation JSON status is not PASS: $($validationObject.status)"
    }

    Write-Host '[4/4] Codex AI analysis report'
    $codexCommand = Get-Command codex.cmd -CommandType Application -ErrorAction Stop | Select-Object -First 1
    $prompt = @"
Write a concise Korean-language Markdown report about the Gate 8 result, using only these files:
- Validation JSON: $relativeRun/validation.json
- Unit-test log: $relativeRun/tests.log

Requirements:
1. Do not rerun any tests or commands.
2. Do not modify any source or result file.
3. Use only facts recorded in the two files; do not invent unsupported claims.
4. Clearly separate a section titled "Machine validation results" from a section titled "AI judgment and interpretation" (translate both headings into Korean).
5. Summarize the test count and outcome, plus the JSON status and important metrics.
6. Include a final conclusion and any remaining limitations or cautions.
"@
    $codex = Invoke-CapturedProcess $env:ComSpec @(
        '/d', '/s', '/c', 'call', $codexCommand.Source,
        'exec', '-s', 'read-only', '--skip-git-repo-check', '-o', $reportRelative, '-'
    ) $prompt
    Write-StageLog (Join-Path $runDirectory 'codex.log') $codex
    if ($codex.ExitCode -ne 0) {
        throw "Codex report generation failed (exit code $($codex.ExitCode)); prior artifacts are preserved."
    }
    if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
        throw 'Codex exited successfully but did not create the Markdown report.'
    }

    Write-Host "Completed: $runDirectory"
    exit 0
}
catch {
    $message = $_.Exception.Message
    if ($null -ne $runDirectory -and (Test-Path -LiteralPath $runDirectory -PathType Container)) {
        $failurePath = Join-Path $runDirectory 'failure.log'
        [System.IO.File]::WriteAllText($failurePath, $message + "`r`n", $script:Utf8NoBom)
        Write-Error "$message`nRun directory and logs were preserved: $runDirectory"
    }
    else {
        Write-Error $message
    }
    exit 1
}
