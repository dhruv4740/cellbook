# Dumps real Excel's ribbon (tab > group > control names) through UI Automation to out/excel_ribbon.json.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$null = $xl.Workbooks.Add()
Start-Sleep 3
$root = [Windows.Automation.AutomationElement]::RootElement
$cond = New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty, 0)
$win = $null
foreach ($w in $root.FindAll('Children', [Windows.Automation.Condition]::TrueCondition)) {
  if ($w.Current.ClassName -eq 'XLMAIN') { $win = $w; break }
}
if (-not $win) { throw 'Excel window not found' }
$tabs = $win.FindAll('Descendants', (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)))
$out = [ordered]@{}
foreach ($t in $tabs) {
  $name = $t.Current.Name
  if (-not $name -or $t.Current.IsOffscreen) { continue }
  try { $t.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { continue }
  Start-Sleep -Milliseconds 700
  $items = @()
  foreach ($e in $win.FindAll('Descendants', [Windows.Automation.Condition]::TrueCondition)) {
    $ct = $e.Current.ControlType.ProgrammaticName
    if ($e.Current.IsOffscreen -or -not $e.Current.Name) { continue }
    if ($ct -in 'ControlType.Button', 'ControlType.SplitButton', 'ControlType.MenuItem', 'ControlType.ComboBox', 'ControlType.CheckBox', 'ControlType.Group') {
      $items += [ordered]@{ type = $ct -replace 'ControlType\.', ''; name = $e.Current.Name; id = $e.Current.AutomationId }
    }
  }
  $out[$name] = $items
  Write-Host "$name : $($items.Count)"
}
New-Item -ItemType Directory -Force scripts/excel-oracle/out | Out-Null
$out | ConvertTo-Json -Depth 4 | Set-Content scripts/excel-oracle/out/excel_ribbon.json -Encoding utf8
$xl.DisplayAlerts = $false
$xl.Quit()
