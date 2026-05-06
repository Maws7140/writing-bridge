param(
    [switch]$SelectOnly,
    [string]$CaptureRectJson = "",
    [string]$OutputFile = ""
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class NativeDpi {
    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    public static void Enable() {
        // PER_MONITOR_AWARE_V2 keeps WinForms mouse coordinates aligned with CopyFromScreen pixels.
        if (!SetProcessDpiAwarenessContext(new IntPtr(-4))) {
            SetProcessDPIAware();
        }
    }
}

public class SmoothOverlayForm : Form {
    public SmoothOverlayForm() {
        DoubleBuffered = true;
        SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.UserPaint,
            true
        );
        UpdateStyles();
    }
}
"@

[NativeDpi]::Enable()
[System.Windows.Forms.Application]::EnableVisualStyles()

function Copy-ScreenRectangle {
    param(
        [System.Drawing.Rectangle]$ScreenRect,
        [string]$Path = ""
    )

    $bitmap = New-Object System.Drawing.Bitmap $ScreenRect.Width, $ScreenRect.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen(
            [System.Drawing.Point]::new($ScreenRect.X, $ScreenRect.Y),
            [System.Drawing.Point]::Empty,
            $ScreenRect.Size
        )

        if ($Path -ne "") {
            $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        } else {
            [System.Windows.Forms.Clipboard]::SetImage($bitmap)
        }
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

if ($CaptureRectJson -ne "") {
    $rectData = $CaptureRectJson | ConvertFrom-Json
    $screenRect = [System.Drawing.Rectangle]::new(
        [int]$rectData.x,
        [int]$rectData.y,
        [int]$rectData.width,
        [int]$rectData.height
    )

    Copy-ScreenRectangle $screenRect $OutputFile
    exit 0
}

$bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
$form = New-Object SmoothOverlayForm
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$form.Bounds = $bounds
$form.TopMost = $true
$form.ShowInTaskbar = $false
$form.BackColor = [System.Drawing.Color]::Black
$form.Opacity = 0.22
$form.Cursor = [System.Windows.Forms.Cursors]::Cross
$form.KeyPreview = $true

$script:startPoint = $null
$script:currentPoint = $null
$script:lastRect = [System.Drawing.Rectangle]::Empty
$script:selectionComplete = $false
$script:screenRect = $null

function Get-SelectionRectangle {
    if ($null -eq $script:startPoint -or $null -eq $script:currentPoint) {
        return [System.Drawing.Rectangle]::Empty
    }

    $x = [Math]::Min($script:startPoint.X, $script:currentPoint.X)
    $y = [Math]::Min($script:startPoint.Y, $script:currentPoint.Y)
    $w = [Math]::Abs($script:startPoint.X - $script:currentPoint.X)
    $h = [Math]::Abs($script:startPoint.Y - $script:currentPoint.Y)
    return [System.Drawing.Rectangle]::new($x, $y, $w, $h)
}

function Invalidate-SelectionRectangle {
    param([System.Drawing.Rectangle]$nextRect)

    $dirty = $nextRect
    if (-not $script:lastRect.IsEmpty) {
        $dirty = [System.Drawing.Rectangle]::Union($dirty, $script:lastRect)
    }

    if (-not $dirty.IsEmpty) {
        $dirty.Inflate(8, 8)
        $form.Invalidate($dirty, $false)
    }

    $script:lastRect = $nextRect
}

$form.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    }
})

$form.Add_MouseDown({
    if ($_.Button -ne [System.Windows.Forms.MouseButtons]::Left) {
        return
    }

    $script:startPoint = $_.Location
    $script:currentPoint = $_.Location
    $script:lastRect = [System.Drawing.Rectangle]::Empty
    Invalidate-SelectionRectangle (Get-SelectionRectangle)
})

$form.Add_MouseMove({
    if ($null -eq $script:startPoint) {
        return
    }

    $script:currentPoint = $_.Location
    Invalidate-SelectionRectangle (Get-SelectionRectangle)
})

$form.Add_MouseUp({
    if ($_.Button -ne [System.Windows.Forms.MouseButtons]::Left -or $null -eq $script:startPoint) {
        return
    }

    $script:currentPoint = $_.Location
    $rect = Get-SelectionRectangle
    if ($rect.Width -lt 4 -or $rect.Height -lt 4) {
        $script:startPoint = $null
        $script:currentPoint = $null
        $script:lastRect = [System.Drawing.Rectangle]::Empty
        $form.Invalidate()
        return
    }

    $script:screenRect = [System.Drawing.Rectangle]::new(
        $bounds.X + $rect.X,
        $bounds.Y + $rect.Y,
        $rect.Width,
        $rect.Height
    )
    $script:selectionComplete = $true
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

$form.Add_Paint({
    $rect = Get-SelectionRectangle
    if ($rect.IsEmpty) {
        return
    }

    $graphics = $_.Graphics
    $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighSpeed
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
    $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 255, 255, 255), 2)
    $fill = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(70, 0, 120, 215))
    try {
        $graphics.FillRectangle($fill, $rect)
        $graphics.DrawRectangle($pen, $rect)
    } finally {
        $pen.Dispose()
        $fill.Dispose()
    }
})

$result = $form.ShowDialog()
if ($result -ne [System.Windows.Forms.DialogResult]::OK -or -not $script:selectionComplete -or $null -eq $script:screenRect) {
    exit 2
}

$form.Dispose()
[System.Threading.Thread]::Sleep(120)

if ($SelectOnly) {
    @{
        x = $script:screenRect.X
        y = $script:screenRect.Y
        width = $script:screenRect.Width
        height = $script:screenRect.Height
    } | ConvertTo-Json -Compress
    exit 0
}

Copy-ScreenRectangle $script:screenRect

exit 0
