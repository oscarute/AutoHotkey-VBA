<# : batch section
@echo off
set "PPTX=%~1"
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ([IO.File]::ReadAllText('%~f0'))"
exit /b
#>
# ---- Extract PowerPoint speaker notes ----
# Usage: double-click (uses the active presentation) or drag a .pptx onto this .bat
try {
    $path = $env:PPTX
    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app = $null }

    if ($path) {
        if (-not $app) { $app = New-Object -ComObject PowerPoint.Application }
        $pres = $app.Presentations.Open($path, -1, 0, 0)   # read-only, no window
        $closeAfter = $true
    } else {
        if (-not $app -or $app.Presentations.Count -eq 0) {
            Write-Host 'No presentation is open. Open one or drag a .pptx onto this file.'
            Read-Host 'Press Enter to exit'; exit
        }
        $pres = $app.ActivePresentation
        $closeAfter = $false
    }

    $sb = New-Object System.Text.StringBuilder
    foreach ($s in $pres.Slides) {
        $notes = ''
        foreach ($sh in $s.NotesPage.Shapes) {
            # 14 = msoPlaceholder, 2 = ppPlaceholderBody (the notes box)
            if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 2 -and $sh.HasTextFrame) {
                $notes = $sh.TextFrame.TextRange.Text
            }
        }
        if (-not $notes) { $notes = '(no notes)' }
        $notes = $notes -replace "`r", "`r`n" -replace "`v", "`r`n"
        [void]$sb.AppendLine("=== Slide $($s.SlideIndex) ===")
        [void]$sb.AppendLine($notes)
        [void]$sb.AppendLine()
    }

    # Output folder: the presentation's folder, or the Desktop if unsaved or stored in OneDrive/SharePoint
    $dir = $pres.Path
    if (-not $dir -or $dir -like 'http*') { $dir = [Environment]::GetFolderPath('Desktop') }
    $name = [IO.Path]::GetFileNameWithoutExtension($pres.Name)
    $out = Join-Path $dir ("Notes_" + $name + ".txt")
    [IO.File]::WriteAllText($out, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    if ($closeAfter) {
        $pres.Close()
        if ($app.Presentations.Count -eq 0) { $app.Quit() }
    }
    Start-Process notepad.exe -ArgumentList "`"$out`""
}
catch {
    Write-Host "Error: $($_.Exception.Message)"
    Read-Host 'Press Enter to exit'
}
