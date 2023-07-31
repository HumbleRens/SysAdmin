Start-Process powershell -Verb RunAs
$eventIDs = @(1151, 1116, 1149, 4699, 4698, 46, 984, 699, 7000, 7009, 4728, 29, 4756, 57, 4732, 33, 4742, 4738, 4723, 4724, 47, 284, 756, 4722, 4741, 4743, 4726, 4720, 4781, 15457, 18457, 19, 20, 21, 4622, 150, 770, 4673, 4717, 4704, 7045, 4697, 6416, 808, 354, 321, 104, 1102, 4616, 4964, 60, 5137, 5141, 5143, 5124, 5136, 5007, 4719, 4739, 4908, 2003, 4950, 2004, 4663, 325, 327, 4794, 472, 847, 564, 732, 4771, 33205, 70, 5382, 4768, 6004, 4625, 131, 4799, 4662, 600, 4769, 4661, 4825, 4648, 5140, 5145, 5142, 4674, 4656, 10, 4624, 13, 800, 4103, 4104, 4688, 11)

# Create a new Excel workbook and worksheet
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Add()

# Set the column headers in the worksheet
$worksheet.Cells.Item(1,1) = "Event ID"
$worksheet.Cells.Item(1,2) = "Time Generated"
$worksheet.Cells.Item(1,3) = "Message"

$row = 2

foreach ($id in $eventIDs) {
    $events = Get-EventLog -LogName Security -InstanceId $id -Newest 30
    
    if ($events) {
        Write-Host "Event ID $id found in Security log:"
        
        foreach ($event in $events) {
            # Write the event details to the worksheet
            $worksheet.Cells.Item($row,1) = $id
            $worksheet.Cells.Item($row,2) = $event.TimeGenerated
            $worksheet.Cells.Item($row,3) = $event.Message
            $row++
        }
    } else {
        Write-Host "Event ID $id not found in Security log."
    }
}

# Autofit the columns
$range = $worksheet.UsedRange
$range.EntireColumn.AutoFit()

# Save and close the workbook
$workbook.SaveAs("$HOME\Desktop\LogFilteredDump.xlsx")
$excel.Quit()
