<#
.SYNOPSIS
    Export tags historian to CSV

.DESCRIPTION
    Is used DvCHDump from "C:\DeltaV\DVUtilities\DvCHDump.exe

.AUTHOR
    https://github.com/xabrs

#>

$programPath = "C:\DeltaV\DVUtilities\DvCHDump.exe"
$dtFormat= "yyyy/MM/dd HH:mm:ss"

Add-Type -AssemblyName System.Windows.Forms
function StartGUI(){
  $form = New-Object System.Windows.Forms.Form
  $form.Width = 700
  $form.Height = 330
  $form.Text = "DeltaVCH2CSV (DvCHDump GUI)"

  $startDateLabel = New-Object System.Windows.Forms.Label
  $startDateLabel.Text = "Start time (UTC):"
  $startDateLabel.Top = 20
  $startDateLabel.Left = 20
  $form.Controls.Add($startDateLabel)

  $startDate = New-Object System.Windows.Forms.DateTimePicker
  $startDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
  $startDate.Top = 43
  $startDate.Left = 20
  $startDate.Width = 150
  $startDate.CustomFormat = $dtFormat
  $form.Controls.Add($startDate)


  $endDateLabel = New-Object System.Windows.Forms.Label
  $endDateLabel.Text = "End time (UTC):"
  $endDateLabel.Top = 80
  $endDateLabel.Left = 20
  $form.Controls.Add($endDateLabel)

  $endDate = New-Object System.Windows.Forms.DateTimePicker
  $endDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
  $endDate.Top = 103
  $endDate.Left = 20
  $endDate.Width = 150
  $endDate.CustomFormat = $dtFormat
  $form.Controls.Add($endDate)

  $intervalLabel = New-Object System.Windows.Forms.Label
  $intervalLabel.Text = "Interval:"
  $intervalLabel.Top = 140
  $intervalLabel.Left = 20
  $form.Controls.Add($intervalLabel)

  $interval = New-Object System.Windows.Forms.TextBox
  $interval.Top = 163
  $interval.Left = 20
  $interval.Width = 150
  $interval.Text = "1H"
  $form.Controls.Add($interval)
  $tooltip = New-Object System.Windows.Forms.ToolTip
  $tooltip.SetToolTip($interval, "Example: 5S, 2M, 1H, 1D,...")

  $tagNameLabel = New-Object System.Windows.Forms.Label
  $tagNameLabel.Text = "Tags:"
  $tagNameLabel.Top = 20
  $tagNameLabel.Left = 200
  $form.Controls.Add($tagNameLabel)


  $tagName = New-Object System.Windows.Forms.TextBox
  $tagName.Top = 43
  $tagName.Left = 200
  $tagName.Width = 450
  $tagName.Height = 220
  $tagName.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
  $tagName.Multiline = $true
  $form.Controls.Add($tagName)

  $checkBox1 = New-Object System.Windows.Forms.CheckBox
  $checkBox1.Text = "Average"
  $checkBox1.Location = New-Object System.Drawing.Point(20, 200)
  $checkBox1.Size = New-Object System.Drawing.Size(90, 30)
  $checkBox1.Checked = $true
  $form.Controls.Add($checkBox1)

  $checkBox2 = New-Object System.Windows.Forms.CheckBox
  $checkBox2.Text = "Interpolated"
  $checkBox2.Location = New-Object System.Drawing.Point(20, 230)
  $checkBox2.Size = New-Object System.Drawing.Size(90, 30)
  $form.Controls.Add($checkBox2)

  $startButton = New-Object System.Windows.Forms.Button
  $startButton.Text = "Export"
  $startButton.Top = 205
  $startButton.Left = 110
  $form.Controls.Add($startButton)
  
  $startButton.Add_Click({
    $resultArray = @()
    $sDate = $startDate.Value.ToString($dtFormat, [System.Globalization.CultureInfo]::InvariantCulture)
    $eDate = $endDate.Value.ToString($dtFormat, [System.Globalization.CultureInfo]::InvariantCulture)
    $outFileName = $startDate.Value.ToString("yyyyMMddHHmmss") +'_'+$endDate.Value.ToString("yyyyMMddHHmmss")+'.csv'
    if (-not $checkBox1.Checked -and -not $checkBox2.Checked){
      [System.Windows.Forms.MessageBox]::Show("No selected checkbox")
      return
    }
    foreach ($tag in $tagName.Text.Split("`n")){
      $arguments = '"PROCESSED AGGREGATE VALUES" '
      if ([string]::IsNullOrEmpty($tag.Trim())) {
        continue
      }
      $arguments += $tag.Trim()
      if ($checkBox1.Checked -and $checkBox2.Checked) {
        $arguments += ' AVE,INTERPOLATED'  
      } elseif ($checkBox1.Checked) {
        $arguments += ' AVE'
      } elseif ($checkBox2.Checked) {
        $arguments += ' INTERPOLATED'
      }
      # available values MIN,MAX,FIRST,LAST - It's not what you think. It comes from the structure of the database.
      # see DVCHDump /?
      
      $arguments += ' "'+$sDate+'"'
      $arguments += ' "'+$eDate+'"'
      $arguments += ' '+$interval.Text.Trim()

      #[System.Windows.Forms.MessageBox]::Show($arguments)
      #return
      $output = RunDVCHDump -arguments $arguments

      if ([string]::IsNullOrEmpty($output)) {
        continue
      }

      # merge the result on the right with the previous results
      $array = ParseResult -lines $output.Split("`n")
      if ($resultArray.Count -eq 0){
        $resultArray = $array
      } else {
        for($i=0;$i -lt $resultArray.Count;$i++){
          if ($resultArray[$i][0] -eq $array[$i][0]){
            $resultArray[$i]+=$array[$i][2..($array[$i].Length - 1)]
          }
        }
      }

    }
    
    
    # Save
    $resultArray | ForEach-Object { $_ -join "`t" } | Out-File -FilePath $outFileName
    [System.Windows.Forms.MessageBox]::Show("Done!")
  })
  $form.ShowDialog()
}

function ParseResult(){ 
  # Return values array with headers from DVCHDump output lines
  param (
   [string[]]$lines
  )
  $array = @()
  $array +=,@('index','datetime')
  $headerIndex = -1
  $tag = ''
  for ($i=0;$i -lt $lines.Count;$i++){
    $line = $lines[$i]
    if ($line -match "'DeltaV=.+ (.+)'"){
      $tag = $matches[1]
    }
    if ($line.IndexOf('     -') -ne -1) {
      $headerIndex = $i
      break
    }
  }
  if ($headerIndex -eq -1){ return }
  $line = $lines[$headerIndex-1]
  $x = $lines[$headerIndex].IndexOf(" -")
  $cols_index = @($x+1)
  while ($x -ne -1){
   $x = $lines[$headerIndex].IndexOf(" -",$x+1)
   if ($x -ne -1) {
     $array[0] += @($tag+" "+$line.Substring($cols_index[-1],$x+1-$cols_index[-1]).Trim().Replace('(timwghtd)AVE','AVG'))
     $cols_index += @($x+1)
   }
  }
  
  $x = $line.Length
  $array[0] += @($tag+" "+$line.Substring($cols_index[-1],$x-$cols_index[-1]).Trim().Replace('(timwghtd)AVE','AVG'))

  foreach ($line in $lines) {
    if ($line -match "^\s*(\d+)\s+([\d\.]+ [\d:]+)\s+") {
        $number = $matches[1]
        $datetime = $matches[2]
        $array += ,@($number, $datetime)

        for ($i=0; $i -lt $cols_index.Count -1;$i++){
            $array[-1] += $line.Substring($cols_index[$i],$cols_index[$i+1]-$cols_index[$i]).Trim()
        }
        $array[-1] += $line.Substring($cols_index[$i],$line.Length-$cols_index[$i]).Trim()
        
    }
  }
  return $array
}

function RunDVCHDump([string] $arguments){
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $process.StartInfo.FileName = $programPath
    $process.StartInfo.Arguments = $arguments
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true

    $process.Start() | Out-Null
    
    $output = $process.StandardOutput.ReadToEnd()
    $errors = $process.StandardError.ReadToEnd()

    $process.WaitForExit()
    
    if ($process.ExitCode -eq 0) {
        return $output
    } else {
        [System.Windows.Forms.MessageBox]::Show("Error! Arguments: $($arguments)`n Code: $($process.ExitCode)`nOut: $errors")
    }
}


StartGUI
