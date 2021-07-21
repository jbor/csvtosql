##############################################################################
##
## CSV to SQL
##
## Tool to support converting csv to sql script.
##
## 14-11-2019 JBOR: Initial
## 09-12-2019 JBOR: Predict-Type added and minor improvements
## 20-07-2021 JBOR: Added a checkbox for PKEY
##
##############################################################################

Function Get-FileName($initialDirectory)
{  
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = "All files (*.*)| *.*"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.filename
}

Function ProcessName ($n)
{
$combo = ($mainForm.Controls | Where-Object {$_.Name -eq ($n + "Combo")})
$type = $combo.SelectedItem
if($type -ne $dropDownArray[2]) {
$n = $n + ", "
} else {
$n = ""
}
return $n
}

Function GetKeyLine($l)
{
$checkBoxes = ($mainForm.Controls | Where-Object { $_.GetType().name -eq "CheckBox" })
$keyCombos = $checkBoxes | Where-Object {$_.Checked -eq $true}
$keyLine = ""
if ($keyCombos.Count -gt 0)
{
foreach ($k in $keyCombos)
{

$key = $k.Name.Substring(0, $k.Name.Length-3)
$combo = ($mainForm.Controls | Where-Object { $_.Name -eq $key+"Combo" })

if($combo.SelectedItem -eq $dropDownArray[0])
{
$keyLine += $key + " = '" + $line.$key + "' AND "
}

if($combo.SelectedItem -eq $dropDownArray[1])
{
$keyLine += $key + " = " + $line.$key + " AND "
}
}

if ($keyLine -ne "")
{
$keyLine = $keyLine.Substring(0,$keyLine.Length-5)
}
}

return $keyLine
}


Function ProcessValue($n, $v)
{
$combo = ($mainForm.Controls | Where-Object {$_.Name -eq ($n + "Combo")})
$type = $combo.SelectedItem
$v = $v -replace "'","''"
if ($v -eq "NULL" -or $type -eq $dropDownArray[1]) {
if ($v.Trim() -eq "") { $v = "NULL" }
$v = $v + ", "
} elseif($type -eq $dropDownArray[0]) {
$v = "'" + $v + "', "
} elseif($type -eq $dropDownArray[2]) {
$v = ""
}
return $v
}

Function Create-UpdateScript()
{
"" | Out-File $outputFile
foreach ($line in $csv) {
$properties = ""
$values = ""
$line | Get-Member -MemberType NoteProperty | ForEach-Object {
$property = (ProcessName($_.Name))
if ($property -ne "") {
$property = $property.Substring(0,$property.Length-2)
$values += $property + " = " + (ProcessValue $_.Name $line.($_.Name))
}
}
$values = $values.Substring(0,$values.Length-2)
$keyLine = GetKeyLine($line)
("UPDATE {0} SET {1} WHERE {2};" -f $tableTextBox.Text, $values, $keyLine) | Out-File -Append $outputFile
}
}

Function Create-InsertScript()
{
"" | Out-File $outputFile
  foreach ($line in $csv) {
$properties = ""
$values = ""
$line | Get-Member -MemberType NoteProperty | ForEach-Object {
$properties += (ProcessName($_.Name))
$values += (ProcessValue $_.Name $line.($_.Name))
}
$properties = $properties.Substring(0,$properties.Length-2)
$values = $values.Substring(0,$values.Length-2)
("INSERT INTO {0} ({1}) VALUES ({2});" -f $tableTextBox.Text, $properties, $values) | Out-File -Append $outputFile
}
}

Function Predict-Type($object, $property)
{
$value = $object.$property
[decimal]$testDecimal = $null
[datetime]$testDate = Get-Date
$isDate = ([datetime]::TryParse($value , [ref]$testDate))
$isDecimal = [decimal]::TryParse($value , [ref]$testDecimal)
if ($isDate -or $isDecimal -or $value.Trim() -eq "") {
return $dropDownArray[1]
} else {
return $dropDownArray[0]
}
}

Function DrawLine($height)
{
$lineLabel = New-Object System.Windows.Forms.Label
$lineLabel.AutoSize = $True
$lineLabel.Text = "---------------------------------------------------------------------------------------------------------"
$lineLabel.Location = new-object System.Drawing.Size(30,(100 + $height))
return $lineLabel
}

Function GetDelimiter()
{
$delimiterArray = ",", ";"
$headerLine = Get-Content $file -First 1
foreach($delimiter in $delimiterArray)
{
if ($headerLine.contains($delimiter))
{
return $delimiter;
}
}
return $delimiterArray[0];
}

$outputFile = "output.sql"
$dropDownArray = "Quoted (string/date)", "Not Quoted (Int/Bool)", "Leave Out"

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

$file = Get-FileName
$csv = import-csv $file -Delimiter (GetDelimiter)

$mainForm = New-Object system.Windows.Forms.Form
$mainForm.text = "CSV to SQL Converter"
$mainForm.BackColor = "#ffffff"
$mainForm.AutoScroll = $True

$propertyLabel = New-Object System.Windows.Forms.Label
$propertyLabel.AutoSize = $True
$propertyLabel.Text = "Generate script to:"
$propertyLabel.Location = new-object System.Drawing.Size(30,20)
$mainForm.Controls.Add($propertyLabel)

$insertButton = new-object System.Windows.Forms.Button
$insertButton.Location = new-object System.Drawing.Size(30,40)
$insertButton.Size = new-object System.Drawing.Size(80,20)
$insertButton.BackColor ="LightGray"
$insertButton.Text = "INSERT"
$insertButton.Add_Click({Create-InsertScript})
$mainForm.Controls.Add($insertButton)

$updateButton = new-object System.Windows.Forms.Button
$updateButton.Location = new-object System.Drawing.Size(150,40)
$updateButton.Size = new-object System.Drawing.Size(80,20)
$updateButton.BackColor ="LightGray"
$updateButton.Text = "UPDATE"
$updateButton.Add_Click({Create-UpdateScript})
$mainForm.Controls.Add($updateButton)

$tableLabel = New-Object System.Windows.Forms.Label
$tableLabel.AutoSize = $True
$tableLabel.Location = new-object System.Drawing.Size(30,70)
$tableLabel.ForeColor = "DarkBlue"
$tableLabel.Text = "Table Name:"
$mainForm.Controls.Add($tableLabel)

$tableTextBox = New-Object System.Windows.Forms.TextBox
$tableTextBox.Location = new-object System.Drawing.Size(150,70)
$tableTextBox.Size = new-object System.Drawing.Size(100,20)
$mainForm.Controls.Add($tableTextBox)

$size = 0

$csv | Get-Member -MemberType NoteProperty | ForEach-Object {
$mainForm.Controls.Add((DrawLine($size)))

$propertyLabel = New-Object System.Windows.Forms.Label
$propertyLabel.AutoSize = $True
$propertyLabel.Text = $_.Name
$propertyLabel.Location = new-object System.Drawing.Size(30,(120 + $size))
$mainForm.Controls.Add($propertyLabel)

$combo = new-object System.Windows.Forms.ComboBox
$combo.Location = new-object System.Drawing.Size(350,(120 + $size))
$combo.Size = new-object System.Drawing.Size(130,30)
$combo.Name = $_.Name + "Combo"

foreach ($item in $dropDownArray) {
$combo.Items.Add($item)
}
$combo.SelectedItem = Predict-Type $csv[0] $_.Name
$mainForm.Controls.Add($combo)

$keyCheckbox = new-object System.Windows.Forms.CheckBox
$keyCheckbox.Location = new-object System.Drawing.Size(500,(120 + $size))
$keyCheckbox.Name = $_.Name + "Key"
$keyCheckbox.Size = new-object System.Drawing.Size(100,20)
    $keyCheckbox.Text = "Primary Key"
    $keyCheckbox.Checked = $false
$mainForm.Controls.Add($keyCheckbox)

$size += 40
}

$mainForm.Controls.Add((DrawLine($size)))
$mainForm.Size  = new-object System.Drawing.Size(1200,(800))

[void]$mainForm.ShowDialog()