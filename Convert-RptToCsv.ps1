<#
.SYNOPSIS
    Converts SQL Server Management Studio (SSMS) .rpt output files to .csv files.

.DESCRIPTION
    SSMS "Results to File"/"Results to Text" output (.rpt) is normally a fixed-width
    grid report:

        Col1       Col2        Col3
        ---------- ----------- ----------
        value1     value2      value3
        value1b    value2b     value3b

        (2 rows affected)

    A single .rpt file can contain several result sets (one per query/grid), each
    separated by a blank line and its own header/dashes block, followed by a
    "(N rows affected)"/"(N row(s) affected)" footer or a "Completion time: ..."
    line (as produced by sqlcmd-style exports).

    This script streams the file line-by-line (it never loads the whole file
    into memory), using the dashed separator line to determine column
    boundaries, and writes each result set to a proper, quoted CSV file as it
    goes. This keeps memory usage low even for multi-gigabyte .rpt files. As a
    fallback, if a file does not look like the fixed-width grid format (e.g. it
    was exported as tab- or comma-delimited text instead), it is converted
    using the detected delimiter.

.PARAMETER Path
    Path to a single .rpt file, or a folder containing .rpt files.

.PARAMETER OutputFolder
    Folder to write the resulting .csv file(s) to. Defaults to the same folder
    as each source file.

.PARAMETER Recurse
    When Path is a folder, also search subfolders for .rpt files.

.PARAMETER Delimiter
    Delimiter to use in the output CSV. Defaults to a comma.

.PARAMETER Encoding
    Encoding used to read the .rpt file(s). Defaults to UTF8 (SSMS commonly
    writes Unicode/UTF-8 output files). Use e.g. 'Default' or 'UTF7' if needed.

.EXAMPLE
    .\Convert-RptToCsv.ps1 -Path .\results.rpt

.EXAMPLE
    .\Convert-RptToCsv.ps1 -Path C:\Exports -OutputFolder C:\Exports\csv -Recurse

.NOTES
    If a .rpt file contains multiple result sets, additional CSVs are created
    with a "_2", "_3", ... suffix before the extension.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Parameter(Position = 1)]
    [string]$OutputFolder,

    [switch]$Recurse,

    [string]$Delimiter = ',',

    [string]$Encoding = 'UTF8'
)

function Get-RptFiles {
    param([string]$InputPath, [switch]$Recurse)

    if (Test-Path -LiteralPath $InputPath -PathType Leaf) {
        return @(Get-Item -LiteralPath $InputPath)
    }
    elseif (Test-Path -LiteralPath $InputPath -PathType Container) {
        $gciParams = @{ Path = $InputPath; Filter = '*.rpt'; File = $true }
        if ($Recurse) { $gciParams['Recurse'] = $true }
        return @(Get-ChildItem @gciParams)
    }
    else {
        throw "Path not found: $InputPath"
    }
}

function Test-DashSeparatorLine {
    param([string]$Line)
    # A column-separator line consists of runs of dashes separated by single spaces,
    # e.g. "---------- ----------- ----------"
    if ([string]::IsNullOrWhiteSpace($Line)) { return $false }
    return $Line -match '^-+( -+)*\s*$'
}

function Get-ColumnSpans {
    param([string]$DashLine)
    # Returns an array of [start,length] pairs based on runs of dashes.
    $spans = @()
    $i = 0
    $len = $DashLine.Length
    while ($i -lt $len) {
        if ($DashLine[$i] -eq '-') {
            $start = $i
            while ($i -lt $len -and $DashLine[$i] -eq '-') { $i++ }
            $spans += , @($start, ($i - $start))
        }
        else {
            $i++
        }
    }
    return $spans
}

function ConvertTo-CsvField {
    param([string]$Value, [string]$Delimiter, [char[]]$QuoteChars)
    if ($null -eq $Value) { return '' }
    if ($Value.Length -eq 0) { return $Value }
    if ($Value.IndexOfAny($QuoteChars) -ge 0 -or $Value.Contains($Delimiter)) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

function Get-OutputPath {
    param([string]$BaseOutFile, [int]$ResultSetIndex)
    if ($ResultSetIndex -le 1) { return $BaseOutFile }
    $dir = [System.IO.Path]::GetDirectoryName($BaseOutFile)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($BaseOutFile)
    $ext = [System.IO.Path]::GetExtension($BaseOutFile)
    return (Join-Path $dir "$name`_$ResultSetIndex$ext")
}

function Get-DotNetEncoding {
    param([string]$Name)
    # Map common PowerShell -Encoding names (which .NET's Encoding.GetEncoding
    # does not accept directly, e.g. 'UTF8') to .NET Encoding instances.
    switch ($Name.ToUpperInvariant()) {
        'UTF8'           { return New-Object System.Text.UTF8Encoding($false) }
        'UTF7'           { return [System.Text.Encoding]::UTF7 }
        'UTF32'          { return [System.Text.Encoding]::UTF32 }
        'ASCII'          { return [System.Text.Encoding]::ASCII }
        'UNICODE'        { return [System.Text.Encoding]::Unicode }
        'BIGENDIANUNICODE' { return [System.Text.Encoding]::BigEndianUnicode }
        'DEFAULT'        { return [System.Text.Encoding]::Default }
        'OEM'            { return [System.Text.Encoding]::Default }
        default          { return [System.Text.Encoding]::GetEncoding($Name) }
    }
}

function Test-FooterLine {
    param([string]$Line)
    # Trailing marker / informational lines seen in SSMS/sqlcmd-style .rpt exports
    # that are not part of the data grid (result-set footers, or warning messages
    # SQL Server sometimes interleaves before/after a result set).
    return ($Line -match '^\(\d+ rows? affected\)') -or
           ($Line -match '^Completion time:') -or
           ($Line -match '^Warning:') -or
           ($Line -match '^Msg \d+, Level \d+, State \d+')
}

function Get-ColumnStartsAndLengths {
    param([System.Collections.Generic.List[int[]]]$Spans)
    # Flatten spans into parallel int[] arrays for fast indexed access in the hot row loop.
    $count = $Spans.Count
    $starts = New-Object int[] $count
    $lengths = New-Object int[] $count
    for ($k = 0; $k -lt $count; $k++) {
        $starts[$k] = $Spans[$k][0]
        $lengths[$k] = $Spans[$k][1]
    }
    return , @($starts, $lengths)
}

function Get-FixedWidthHeaders {
    param([string]$Line, [int[]]$Starts, [int[]]$Lengths)
    $lineLen = $Line.Length
    $count = $Starts.Length
    $values = New-Object string[] $count
    for ($k = 0; $k -lt $count; $k++) {
        $start = $Starts[$k]
        if ($start -lt $lineLen) {
            $avail = $lineLen - $start
            $length = $Lengths[$k]
            if ($length -gt $avail) { $length = $avail }
            $values[$k] = $Line.Substring($start, $length).Trim()
        } else {
            $values[$k] = ''
        }
    }
    return $values
}

function Write-HeaderCsvLine {
    param([System.IO.TextWriter]$Writer, [string[]]$Headers, [string]$Delimiter, [char[]]$QuoteChars)
    for ($k = 0; $k -lt $Headers.Length; $k++) {
        if ($k -gt 0) { $Writer.Write($Delimiter) }
        $Writer.Write((ConvertTo-CsvField -Value $Headers[$k] -Delimiter $Delimiter -QuoteChars $QuoteChars))
    }
    $Writer.Write("`r`n")
}

function Write-FixedWidthCsvRow {
    param(
        [System.IO.TextWriter]$Writer,
        [string]$Line,
        [int[]]$Starts,
        [int[]]$Lengths,
        [string]$Delimiter,
        [char[]]$QuoteChars
    )
    $lineLen = $Line.Length
    $count = $Starts.Length
    for ($k = 0; $k -lt $count; $k++) {
        if ($k -gt 0) { $Writer.Write($Delimiter) }
        $start = $Starts[$k]
        if ($start -lt $lineLen) {
            $avail = $lineLen - $start
            $length = $Lengths[$k]
            if ($length -gt $avail) { $length = $avail }
            $value = $Line.Substring($start, $length).Trim()
            if ($value.Length -gt 0) {
                if ($value.IndexOfAny($QuoteChars) -ge 0 -or $value.Contains($Delimiter)) {
                    $Writer.Write('"')
                    $Writer.Write($value.Replace('"', '""'))
                    $Writer.Write('"')
                } else {
                    $Writer.Write($value)
                }
            }
        }
    }
    $Writer.Write("`r`n")
}

function Write-DelimitedCsvRow {
    param(
        [System.IO.TextWriter]$Writer,
        [string[]]$Values,
        [string]$Delimiter,
        [char[]]$QuoteChars
    )
    for ($k = 0; $k -lt $Values.Length; $k++) {
        if ($k -gt 0) { $Writer.Write($Delimiter) }
        $Writer.Write((ConvertTo-CsvField -Value $Values[$k].Trim() -Delimiter $Delimiter -QuoteChars $QuoteChars))
    }
    $Writer.Write("`r`n")
}

function New-Utf8StreamWriter {
    param([string]$Path)
    $enc = New-Object System.Text.UTF8Encoding($true)
    $stream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read, 1MB)
    $writer = New-Object System.IO.StreamWriter($stream, $enc, 1MB)
    return $writer
}

function Convert-RptFileStreaming {
    param([System.IO.FileInfo]$File, [string]$BaseOutFile, [string]$Delimiter, [string]$Encoding)

    $encObj = Get-DotNetEncoding -Name $Encoding
    $fileStream = New-Object System.IO.FileStream($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read, 1MB)
    $reader = New-Object System.IO.StreamReader($fileStream, $encObj, $true, 1MB)
    $converted = $false
    $resultSetCount = 0
    $quoteChars = [char[]]@('"', "`r", "`n")

    try {
        $line = $reader.ReadLine()
        while ($null -ne $line) {
            if ([string]::IsNullOrWhiteSpace($line) -or (Test-FooterLine -Line $line)) {
                $line = $reader.ReadLine()
                continue
            }

            $headerLine = $line
            $dashLine = $reader.ReadLine()

            if ($null -eq $dashLine -or [string]::IsNullOrWhiteSpace($dashLine) -or (Test-FooterLine -Line $dashLine)) {
                # headerLine wasn't actually a header (e.g. a stray SQL Server
                # "Warning:"/"Msg ..." message); drop it as noise and keep
                # scanning from dashLine (which may itself be blank/footer/noise
                # and get skipped on the next loop iteration, or may be a real
                # header).
                $line = $dashLine
                continue
            }

            if (-not (Test-DashSeparatorLine -Line $dashLine)) {
                # headerLine looked like data/noise, not a real header, since the
                # line after it isn't a dash separator. Drop headerLine as noise
                # and retry using dashLine as the new header candidate (without
                # losing it).
                $line = $dashLine
                continue
            }

            $spansRaw = Get-ColumnSpans -DashLine $dashLine
            $spans = New-Object System.Collections.Generic.List[int[]]
            foreach ($s in $spansRaw) { $spans.Add([int[]]@($s[0], $s[1])) }
            $startsLengths = Get-ColumnStartsAndLengths -Spans $spans
            $starts = $startsLengths[0]
            $lengths = $startsLengths[1]

            $headers = Get-FixedWidthHeaders -Line $headerLine -Starts $starts -Lengths $lengths

            $resultSetCount++
            $outFile = Get-OutputPath -BaseOutFile $BaseOutFile -ResultSetIndex $resultSetCount
            $writer = New-Utf8StreamWriter -Path $outFile
            try {
                Write-HeaderCsvLine -Writer $writer -Headers $headers -Delimiter $Delimiter -QuoteChars $quoteChars

                $rowCount = 0
                $line = $reader.ReadLine()
                while ($null -ne $line -and -not [string]::IsNullOrWhiteSpace($line) -and -not (Test-FooterLine -Line $line)) {
                    Write-FixedWidthCsvRow -Writer $writer -Line $line -Starts $starts -Lengths $lengths -Delimiter $Delimiter -QuoteChars $quoteChars
                    $rowCount++
                    if ($rowCount % 100000 -eq 0) {
                        Write-Host "    ...$rowCount rows written so far"
                    }
                    $line = $reader.ReadLine()
                }
            }
            finally {
                $writer.Flush()
                $writer.Close()
            }

            Write-Host "  Wrote $rowCount row(s) -> $outFile"
            $converted = $true

            # Skip blank/footer lines before looking for the next result set.
            while ($null -ne $line -and ([string]::IsNullOrWhiteSpace($line) -or (Test-FooterLine -Line $line))) {
                $line = $reader.ReadLine()
            }
        }
    }
    finally {
        $reader.Close()
    }

    return $converted
}

function Convert-DelimitedRptStreaming {
    param([System.IO.FileInfo]$File, [string]$BaseOutFile, [string]$Delimiter, [string]$Encoding)

    $encObj = Get-DotNetEncoding -Name $Encoding
    $fileStream = New-Object System.IO.FileStream($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read, 1MB)
    $reader = New-Object System.IO.StreamReader($fileStream, $encObj, $true, 1MB)
    $quoteChars = [char[]]@('"', "`r", "`n")
    try {
        $headerLine = $reader.ReadLine()
        while ($null -ne $headerLine -and [string]::IsNullOrWhiteSpace($headerLine)) {
            $headerLine = $reader.ReadLine()
        }
        if ($null -eq $headerLine) { return $false }

        $sourceDelim = if ($headerLine.Contains("`t")) { "`t" }
                       elseif ($headerLine.Contains(',')) { ',' }
                       elseif ($headerLine.Contains(';')) { ';' }
                       else { $null }
        if (-not $sourceDelim) { return $false }

        $splitChars = [char[]]@($sourceDelim)
        $headers = $headerLine.Split($splitChars) | ForEach-Object { $_.Trim() }

        $writer = New-Utf8StreamWriter -Path $BaseOutFile
        $rowCount = 0
        try {
            Write-HeaderCsvLine -Writer $writer -Headers $headers -Delimiter $Delimiter -QuoteChars $quoteChars
            $line = $reader.ReadLine()
            while ($null -ne $line) {
                if ([string]::IsNullOrWhiteSpace($line) -or (Test-FooterLine -Line $line)) {
                    $line = $reader.ReadLine()
                    continue
                }
                $values = $line.Split($splitChars)
                Write-DelimitedCsvRow -Writer $writer -Values $values -Delimiter $Delimiter -QuoteChars $quoteChars
                $rowCount++
                if ($rowCount % 100000 -eq 0) {
                    Write-Host "    ...$rowCount rows written so far"
                }
                $line = $reader.ReadLine()
            }
        }
        finally {
            $writer.Flush()
            $writer.Close()
        }

        Write-Host "  Wrote $rowCount row(s) -> $BaseOutFile (delimiter '$sourceDelim' detected)"
        return $true
    }
    finally {
        $reader.Close()
    }
}

# --- Main ---

$files = Get-RptFiles -InputPath $Path -Recurse:$Recurse
if ($files.Count -eq 0) {
    Write-Warning "No .rpt file(s) found at '$Path'."
    return
}

foreach ($file in $files) {
    $destFolder = if ($OutputFolder) {
        if (-not (Test-Path -LiteralPath $OutputFolder)) {
            New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        }
        $OutputFolder
    } else {
        $file.DirectoryName
    }

    $baseOutFile = Join-Path $destFolder ($file.BaseName + '.csv')
    $sizeMb = [Math]::Round($file.Length / 1MB, 1)
    Write-Host "Converting $($file.FullName) ($sizeMb MB)"

    $ok = Convert-RptFileStreaming -File $file -BaseOutFile $baseOutFile -Delimiter $Delimiter -Encoding $Encoding
    if (-not $ok) {
        $ok = Convert-DelimitedRptStreaming -File $file -BaseOutFile $baseOutFile -Delimiter $Delimiter -Encoding $Encoding
    }
    if (-not $ok) {
        Write-Warning "  Could not detect a known .rpt format for $($file.FullName); file left unconverted."
    }
}
