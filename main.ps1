Function Main() {
  $taskCount = CountTasks
  $title = $(Get-Date -Format "d") + " のタスク"
  $message = "もうすぐ〆切: " + $taskCount[0] + " 個"
  $detail = "そろそろ〆切: " + $taskCount[1] + " 個"
  ShowToast -title $($title) -message $message -detail $detail
}

Function CountTasks {
  $dailyTasksCount = 0
  $weeklyTasksCount = 0

  $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
  $needQuit = $false
  if ($null -eq $outlookProcess) {
    $needQuit = $true
  }

  $outlook = New-Object -ComObject Outlook.Application
  try {
    $namespace = $outlook.GetNamespace("MAPI")
    $tasks = $namespace.GetDefaultFolder(28).Items # https://docs.microsoft.com/ja-jp/office/vba/api/outlook.oldefaultfolders
    $today = Get-Date
    foreach ($task in $tasks) {
      # Write-Host $task.Subject, $task.DueDate
      if ($task.Complete) {
        continue
      }
      $timeSpan = ($task.DueDate - $today).Days
      if ($timeSpan -le 5) {
        $weeklyTasksCount += 1
        if ($timeSpan -le 0) {
          $dailyTasksCount += 1
        }
      }
    }
  } finally {
    if ($needQuit) {
      [void]$outlook.Quit()
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
    }
  }

  return $dailyTasksCount, $weeklyTasksCount
}

# from https://qiita.com/magiclib/items/12e2a9e1e1e823a7fa5c
# see more https://docs.microsoft.com/ja-jp/windows/uwp/design/shell/tiles-and-notifications/adaptive-interactive-toasts
# see about app protocol https://www.codeproject.com/Articles/1187127/Windows-toast-notifications-without-a-COM-server, https://stackoverflow.com/questions/43785421/how-register-a-protocol-for-a-windows-toast
Function ShowToast {
  [CmdletBinding()]
  PARAM (
    [Parameter(Mandatory = $true)][String] $title,
    [Parameter(Mandatory = $false)][String] $message,
    [Parameter(Mandatory = $false)][String] $detail
  )

  [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
  [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
  [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

  # $app_id = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
  $app_id = 'Microsoft.Todos_8wekyb3d8bbwe!App'
  $content = @"
<?xml version="1.0" encoding="utf-8"?>
<toast launch="ms-todo:" activationType="protocol">
  <visual>
      <binding template="ToastGeneric">
          <text>$($title)</text>
          <text>$($message)</text>
          <text>$($detail)</text>
      </binding>
  </visual>
</toast>
"@
  $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
  $xml.LoadXml($content)
  $toast = New-Object Windows.UI.Notifications.ToastNotification $xml
  [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app_id).Show($toast)
}

Main

