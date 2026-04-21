param(
    [Parameter(Mandatory=$true)]
    [string]$DocxPath
)

$word = $null
$doc = $null
$mutex = $null
$mutexAcquired = $false

try {
    $DocxPath = [System.IO.Path]::GetFullPath($DocxPath)

    if (-not (Test-Path $DocxPath)) {
        exit 1
    }

    $pdfPath = [System.IO.Path]::ChangeExtension($DocxPath, ".pdf")

    # Named mutex ensures only one conversion runs at a time.
    $mutex = New-Object System.Threading.Mutex($false, "Global\DocxToPdfConverter")
    $mutexAcquired = $mutex.WaitOne(120000)  # 120s failsafe timeout

    if (-not $mutexAcquired) {
        exit 1
    }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone

    $doc = $word.Documents.Open($DocxPath, $false, $true)  # ConfirmConversions=false, ReadOnly=true

    # wdFormatPDF = 17
    $doc.SaveAs([ref]$pdfPath, [ref]17)

    $doc.Close([ref]0)  # wdDoNotSaveChanges = 0
    $doc = $null

    $word.Quit()
    $word = $null
}
catch {
    # Silent failure by design
}
finally {
    # Release COM objects
    if ($doc) {
        try { $doc.Close([ref]0) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
    }
    if ($word) {
        try { $word.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Wait for all hidden (windowless) Word processes to exit.
    # Since we hold the mutex, no other conversion is running.
    # Any WINWORD with no visible window is either ours or an orphan.
    # Word instances the user has open have a visible window and are safe.
    $waited = 0
    while ($waited -lt 8000) {
        $hidden = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq 0 })
        if ($hidden.Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
        $waited += 500
    }

    # If any hidden Word is STILL alive after 8s, force-kill
    Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq 0 } |
        ForEach-Object {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }

    # Wait for the kills to take effect
    Start-Sleep -Milliseconds 500

    # Release the mutex so the next queued conversion can start
    if ($mutex) {
        if ($mutexAcquired) {
            try { $mutex.ReleaseMutex() } catch {}
        }
        $mutex.Dispose()
    }
}
