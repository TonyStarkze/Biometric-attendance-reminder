Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore

function Minimize-FullScreenApplications {
    $fullScreenApps = Get-WmiObject Win32_Process | Where-Object { 
        $_.Name -eq 'POWERPNT.EXE' -or 
        $_.Name -eq 'EXCEL.EXE' -or 
        $_.Name -eq 'chrome.exe' -or 
        $_.Name -eq 'msedge.exe'
    }

    foreach ($process in $fullScreenApps) {
        if ($process.MainWindowHandle -ne $null) {
            $hwnd = (Add-Type -Name NativeMethods -Namespace Win32 -PassThru -MemberDefinition '
              [DllImport("user32.dll")]
              public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);')

            $hwnd::ShowWindow($process.MainWindowHandle, 6)  # 6 corresponds to minimizing the window
        }
    }
}

function Show-PunchOutWindow {
    Minimize-FullScreenApplications  # Minimize full-screen applications before showing punch-out window

    $window = New-Object System.Windows.Window
    $window.WindowStyle = 'None'
    $window.WindowState = 'Maximized'
    $window.Background = 'Green'
    
    $window.Topmost = $true

    $grid = New-Object System.Windows.Controls.Grid
    $window.Content = $grid

    $textBlockPunchOut = New-Object System.Windows.Controls.TextBlock
    $textBlockPunchOut.Text = "Punch Out Time!"
    $textBlockPunchOut.FontSize = 60
    $textBlockPunchOut.FontWeight = [System.Windows.FontWeights]::Bold
    $textBlockPunchOut.Foreground = 'White'
    $textBlockPunchOut.HorizontalAlignment = 'Center'
    $textBlockPunchOut.VerticalAlignment = 'Center'
    $grid.Children.Add($textBlockPunchOut)

    $verticalSpace = New-Object System.Windows.Controls.StackPanel
    $verticalSpace.Orientation = 'Vertical'
    $verticalSpace.Margin = New-Object System.Windows.Thickness(0, 50, 0, 0)
    $grid.Children.Add($verticalSpace)

    $button = New-Object System.Windows.Controls.Button
    $button.Content = "OKAY BRO!"
    $button.FontSize = 30
    $button.FontWeight = [System.Windows.FontWeights]::Bold
    $button.Width = 500
    $button.Padding = New-Object System.Windows.Thickness(20)
    $button.VerticalAlignment = 'Center'
    $button.HorizontalAlignment = 'Center'
    $button.Add_Click({
        $window.Hide()  # Hide the window instead of closing it
    })
    $button.Margin = New-Object System.Windows.Thickness(0, 500, 0, 0)
    $grid.Children.Add($button)

    $window.ShowDialog()
}

# Hide the PowerShell window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

public static void Hide()
{
    IntPtr hWnd = GetConsoleWindow();
    if (hWnd != IntPtr.Zero)
    {
        ShowWindow(hWnd, 0);
    }
}'

[Console.Window]::Hide()

while ($true) {
    $currentTime = Get-Date
    if ($currentTime.Hour -ge 9 -and $currentTime.Hour -lt 17) {
        $minutesUntilNextPunchOut = 50 - $currentTime.Minute
        if ($minutesUntilNextPunchOut -le 0) {
            $minutesUntilNextPunchOut += 60
        }

        Start-Sleep -Seconds ($minutesUntilNextPunchOut * 60)
        Show-PunchOutWindow
    } else {
        break
    }
}
