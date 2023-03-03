Clear-Host
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

#####################################################################################
############################# Update History ########################################
#####################################################################################

$UpdateHistory =@"

======================================= Start =======================================
    v0.1.0 = Initial version.
    v0.1.1 = Fixed version display bug.
    v0.1.2 = Add feature to delete last record.
	         Separated add case tip.
             Added test option to bulk import data.
    v0.1.3 = Modified prompt message.
    v0.1.4 = Added time before added message.
    v0.1.5 = Shows case note in added message
    v0.1.6 = Remove blank lines from csv when view data.
             Added monthly work history
             Renamed 'TodayWork' to 'DailyWork'.
    v0.1.7 = Fixed a time record bug.
             Added case detail in daily/monthly work review.
    v0.1.8 = Modified prompt code.
    v0.1.9 = Auto detect OneDrive linked profile.
             Selected current month as default month in Monthly history.
    v0.2.0 = Modified Update History code, enhanced performance.
    v0.2.1 = Removed Auto Detect OneDrive.
             Fixed a bug of Monthly history rating.
             Modified some text.
             Add new feature of remove specific case.
             Add new feature of fix csv header.
             Add new feature of modify existing case rating.
             fixed some bug in view case history.
    v0.2.2 = Fixed a bug of check default path.
             Modified some text.
             Fixed a bug of remove specific case.
    v0.2.3 = Modified add new case logic.
    v0.2.4 = Fixed a bug of daily history.
             Fixed a bug of date in monthly history.
    v0.2.5 = Fixed a bug in Daily & monthly history.
    V0.2.6 = Fixed a bug in remove empty line.
    v0.2.7 = Add new feature of fix header missing.
             Add new feature of check case ID format.
    v0.2.8 = Fixed bug of add new case.
    v0.2.9 = Fixed a bug in remove empty line.
    v0.3.0 = Add a feature of debug.
             Trim input message.
    v0.3.1 = Removed prompt in View-AllWork.
    v0.3.2 = Add a feature of case details with GUI.
    v0.3.3 = text modification
    v0.3.4 = You can now add '-' in case id.
    v0.3.5 = Add a feature of auto remove line break.
    v0.3.6 = Fixed a bug of wrong monthly display.
    v0.3.7 = Text modification.
             'Renamed from Daily Work' to 'Case Tracker'.
             Changed version ID, now started with 0.1.
             Modify version check function.
    v0.3.8 = Fixed bug with call wrong function.
             Fixed a bug of not display monthly review correctly.
    v0.3.9 = Fixed a bug when remove case.


    Latest Releases: 
    Releases: https://github.com/lischen2014/My-Case-Tracker

                                                   Author: Leon
                                                   Email: leon2014@vip.qq.com
======================================= End =======================================

"@


#####################################################################################
################################ Functions ##########################################
#####################################################################################

function Show-Menu{
    param (
    [string]$Title = 'Menu'
    )
    Write-Host ""
    Write-Host "================ $Title ================"
    write-host ""
    Write-Host "1: Press '1' to add a case."
    Write-Host "2: Press '2' to remove a case"
    Write-Host "3: Press '3' to change an existing case rating"
    Write-Host "4: Press '4' to view daily history."
    Write-Host "5: Press '5' to view monthly history."
    Write-Host "7: Press '6' to view all history."
    Write-Host "8: Press '8' to open CSV folder via file explorer."
    Write-Host "9: Press '9' to view update log."
    Write-Host "Q: Press 'Q' to quit."
    Write-Host ""
}


function Start-Menu{
	do{
        Refresh-Date
        Remove-EmptyLine
        Check2Fix-Header
        Show-Menu -Title "My Case Tracker $Current_Version"
        write-host "Note. You can add a case id directly from main menu."
        $UserInput = ((Read-Host "Make selection/Input case ID").trim()) -replace "\r?\n", " "
        switch($UserInput)
        {
            '1'{
                Add-Case
                }
            
            '2'{
                Remove-SpecificCase
                }
            
            '3'{
                Change-Rating
            }
            '4'{
                View-DailyWorkWithRating
            }
            '5'{
                View-MonthlyWorkWithRating
            }
            '6'{
                View-AllWork
            }
            '8'{
                # open csv
                explorer $filedir
            }
            '9'{
                $UpdateHistory
                pause
            }
            # end selection
            default{
                Write-Host ""
                if($UserInput.Length -gt 5){
                    Add-Case -userinput $UserInput
                }
                else{
                    Write-Warning "Invalid Option."
                }
            }
        }
	}
	until($UserInput -eq 'Q')
}


function Get-ProfilePath {
    if ($KeepLocal -eq $false){
        $RegEx_EN_OD = "^OneDrive - .{1,}$"
        $UserProfileFolders = (Get-ChildItem "C:\Users\$env:USERNAME").Name

        # Enterprise OneDrive
        if ($UserProfileFolders | ?{$_ -match "$RegEx_EN_OD"}){
            $prefix = $UserProfileFolders | ?{$_ -match "$RegEx_EN_OD"}
        }
        # Personal OneDrive
        elseif ($UserProfileFolders | ?{$_ -like "OneDrive"}){
            $prefix = "OneDrive"
        }

        if($prefix){
            $profilepath = "C:\Users\$env:USERNAME\$prefix"
        }
        else{
            $profilepath = "C:\Users\$env:USERNAME"
        }
    }
    else{
        $profilepath = "C:\Users\$env:USERNAME"
    }

    return $profilepath
}


function Get-LatestVersionId {
    
    $pattern = 'v(\d+(\.\d+){0,3})'
    $matches = [regex]::Matches($UpdateHistory, $pattern)
    
    if ($matches.Count -gt 0) {
        $latestMatch = $matches[$matches.Count - 1]
        return $latestMatch.Groups[0].Value
    }
    
    return $null
}


function Refresh-Date{
    # refresh the $date if script keep running the other day
    $global:date = get-date -format "MM/dd/yyyy"
}


function Write-HostWithTime{
    param (
        [String] $Message,
        [System.ConsoleColor] $ForegroundColor = $Host.UI.RawUI.ForegroundColor
    )

    if ($ForegroundColor -eq '-1'){
        $ForegroundColor = [ConsoleColor]"White"
    }

    $time = Get-Date -Format "HH:mm:ss"
    Write-Host "[$time] $Message" -ForegroundColor $ForegroundColor
}


function Prompt-Confirm{
    # Version: 1.2
    # Author: Leon Jiang

    param(
        [Parameter(mandatory=$false)]
        $action = "continue this action",
        [Parameter(mandatory=$false)]
        $id,
        [Parameter(mandatory=$false)]
        [Bool]$prompt
    )

    # Determine prompt or not
    if($prompt){
        # Return value: 0 for True, 1 for False

        $title = "Confirm '$action'"
    
        if($id){
            $question = "Do you want to '$action' for '$id'?"
        }
        else{
            $question = "Do you want to '$action'?"
        }

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
    
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        # or we can add global variable then remove the return line.
        # $global:decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        if ($decision -eq 0) {
            Write-Host ' Confirmed'
        }
        else{
            Write-Host ' Cancelled'
        }
    }
    else{
        # Return value: 0 for True, 1 for False
        Do {
            $decision = Read-Host -Prompt "Do you want to '$action'? (y/n)"
        }
        Until ($decision -eq 'y' -or $decision -eq 'n')
    }

    return $decision
}


function New-Csv{
    write-host "CSV is not created, creating..."

    # Create csv file 
    $NewExcel = New-Object -ComObject Excel.Application
    $NewExcel.visible = $false
    $NewWorkbook = $NewExcel.workbooks.add() # Add workbook (a file)
    $NewSheet = $NewWorkbook.worksheets.item(1) # Rename workbook sheet 1 as NewSheet
    $NewSheet.name = "MyVolume"
    $NewWorkbook.Styles("Normal").font.name = "Calibri"
    $NewWorkbook.Styles("Normal").font.size = 11
    [void]$NewSheet.Columns.AutoFit()

    # Add subject
    $NewSheet.cells.item(1,1) = '# Do not change the name of columns otherwise script may failed'
    $NewSheet.cells.item(2,1) = 'Date'
    $NewSheet.cells.item(2,2) = 'Time'
    $NewSheet.cells.item(2,3) = 'Case'
    $NewSheet.cells.item(2,4) = 'Rating'

    # Save the file
    try{
        $NewWorkbook.SaveAs("$filenosuffix",[Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV) # xlCSV specifies the CSV file format
        write-host "CSV is created, the path is:"
        write-host $file -ForegroundColor Cyan
        Write-Warning "Do not change the CSV file path/column names, or you have to modify the path/column in scripts!"
    }
    catch{$e}
    finally{
        $NewWorkbook.close($true)
        $NewExcel.quit()
    }
}


function Add-Case{
    param(
        [hashtable]$case=@{},
        $userinput
    )

    $case.date = get-date -format "MM/dd/yyyy"
    $case.time = get-date -Format "HH:mm:ss"

    # Add Case ID
    [Bool]$CasePass = $false

    while (!$CasePass){
        if (!$userinput -eq $false){
            $case.case = $userinput # Pass userinput as caseid if exist
        }
        elseif (!$case.case){
            $case.case = ((read-host "please input case ID").trim()) -replace "\r?\n", " "
        }
        
        # Check if legal
        if ($case.case -match $RegExCaseID){
            [Bool]$CasePass = $true
        }
        else{
            Write-Warning "Case ID can only contains digits, English letters and underline."
            return
        }
    }
    
    # Check Case ID length
    if (($case.case).Length -le 7){
        Write-Host ""
        Write-Warning "Action failed, the length of Case ID should larger than 7."
        return
    }

    # Add Rating
    $case.Rating = (Read-Host "[Optional] Please type case rating").trim()
    while ($case.Rating -eq 0){
        Write-Host "Rating cannot be 0, if you want skip please press Enter key."
        $case.Rating = (Read-Host "[Optional] Please type case rating").trim()
    }

    # Add Note
    $case.content = $case.date, $case.time, $case.case, $case.Rating -join ','
    write-host ""
    try{
        Transmit-Case -case $case
        if(!$case.Rating){
            Write-HostWithTime -Message "Message: $($case.case) is recorded." -ForegroundColor Cyan
        }
        else{
            Write-HostWithTime -Message "Message: $($case.case) - $($case.Rating) is recorded." -ForegroundColor Cyan
        }
    }
    catch{$e}
}


function Transmit-Case{
    param(
        $case
    )
    foreach ($singlecase in $case.content){
        $singlecase | add-content -path $file
    }
}


function View-DailyWorkWithRating{
    Remove-EmptyLine
    $csv = Import-Csv $file | Where-Object {([DateTime]$_.Date) -eq $date} 

    # Parse data
    $TodayWork = $csv
    $Today = [ordered]@{}

    # Calculate how many cases closed
    $Today.CaseDetail = $TodayWork
    $Today.Case = ($Today.CaseDetail | Measure-Object).Count
    
    # Calculate rating
    $Ratings = $TodayWork | Where-Object {$_.Rating -ne ''} | Select-Object -ExpandProperty Rating
    $Today.Rating = ($Ratings | Measure-Object -Average).Average
    $Today.Rating = [math]::round($Today.Rating ,2)

    $TodayReview = New-Object PSObject -property @{
        Date= $date
        Cases= $Today.Case
        AvgRating = $Today.Rating
    }

    # Summary
    $TodayReview | Format-Table -Property Date, Cases, AvgRating
    Display-Details -table $TodayWork
}


function View-MonthlyWorkWithRating{
    Remove-EmptyLine

    # Select time range
    try {
        $SOMUserInput = (Read-Host "Enter a month (Eg: 2023-01)").trim()
        if(!$SOMUserInput){
            [DateTime]$StartOfMonth = (Get-Date).ToString("yyyy-MM") + "-01"
            Write-Host "No input, selected current month"
        }
        else{
            $StartOfMonth = [DateTime]$SOMUserInput
        }
        break
    }
    catch{
        Write-Host "Invalid time, try again." -f Red
    }
    $EndOfMonth = $StartOfMonth.AddMonths(1)

    $csv = Import-Csv $file | Where-Object -FilterScript {([DateTime]::Parse($_."Date") -ge $StartOfMonth) -and ([DateTime]::Parse($_."Date") -lt $EndOfMonth)} 
    
    # Parse data
    $MonthWork = $csv
    $Month = [ordered]@{}

    # Calculate how many cases closed
    $Month.CaseDetail = $MonthWork
    $Month.Case = ($Month.CaseDetail | Measure-Object).Count
    
    # Calculate rating
    $Ratings = $MonthWork | Where-Object {$_.Rating -ne ''} | Select-Object -ExpandProperty Rating
    $Month.Rating = ($ratings | Measure-Object -Average).Average
    $Month.Rating = [math]::round($Month.Rating ,2)

    # Display the current data
    $MonthReview = New-Object PSObject -property @{
        Date= ([String]$StartOfMonth).Substring(3,7)
        Cases= $Month.Case
        AvgRating = $Month.Rating
    }

    # Summary
    $MonthReview | Format-Table -Property Date, Cases, AvgRating
    Display-Details -table $MonthWork
}


function View-AllWork{
    Remove-EmptyLine
    $AllWork = import-csv $File 
    # $AllWork | Format-Table -AutoSize
    $AllWork | Out-GridView
}


function Display-Details{
    param($table)

    write-host ""
    $decision = Prompt-Confirm -action 'display daily work details'
    if ($decision -eq 'y'){
        # $table | Format-Table -AutoSize
        $table | Out-GridView
    }
    else{
        write-host "Cancelled,back to main menu."
    }
}


function Remove-EmptyLine{
    # Remove csv empty lines and txt empty lines
    (gc $file) | ? {($_.trim() -ne ",,,,") -and ($_.trim() -ne "") -and ($_.trim() -ne ",,,,,") } | Set-Content $file
}


function Remove-SpecificCase{

    $FileContent = Get-Content -path $File

    if($FileContent){
        write-host ""

        # user input case id
        [Bool]$CasePass = $false

        while (!$CasePass){
            Write-Warning "You're trying to delete a record."
            $CaseNeedsDelete = (read-host "please input case ID").trim()

            # Check if legal
            if (($CaseNeedsDelete -match $RegExCaseID) -and ($CaseNeedsDelete)){
                [Bool]$CasePass = $true
            }
            else{
                Write-Warning "Case ID can only contains digits, English letters and underline"
                $CaseNeedsDelete = read-host ("please input case ID").trim()
            }
        }

        # search and match the case
        $OldLines = $FileContent -split "`n"
        $NewLines = @()

        
        foreach ($line in $OldLines){
            if (!($line -match $CaseNeedsDelete)){
                $NewLines += $line
            }
            else{
                # no action
            }
        }

        Set-Content -Path $File -value $NewLines
        Write-Host ""
        Write-Host "Removed $CaseNeedsDelete Success" -f Cyan
    }
    else{
        write-host ""
        Write-Host "Message: The CSV does not contain any records."
    }
}


function Change-Rating {
    $Filecontent = Get-Content -path $file

    if($Filecontent){
        # user input a case id
        $CaseNeedsChange = Read-Host "Please input the case ID you need to change rating"
        $CaseRating = Read-Host "Please input the new rating"

        # search and match the case
        $OldLines = $filecontent -split "`n"
        $NewLines = @()

        foreach ($line in $OldLines){
            if ($line -match $CaseNeedsChange){
                $linedetails = $line -split ","
                $linedetails[3] = $CaseRating
                $NewLine = $linedetails -join ","
                $line = $NewLine
            }
            else{
                # no action
            }
            $NewLines += $line
        }

        Set-Content -Path $file -value $NewLines
        Write-Host "Changed $CaseNeedsChange to new rating $CaseRating Success"
    }
    else{
        write-host ""
        Write-Host "Message: The CSV does not contain any records."
    }
}


function Check2Fix-Header{
    $lines = get-content -path $File

    if ($lines[1] -notmatch 'Date'){
        Write-Host "Found CSV header is missing, auto fixed to correct header." -f Cyan

        $DefaultHeaderLine1 = "# Do not change the name of columns otherwise script may failed,,,"

        if ($EnabledRating -eq $true){
            $DefaultHeaderLine2 = "Date,Time,Case,Rating"
        }
        else {
            $DefaultHeaderLine2 = "Date,Time,Case,Type,Note"
        }

        $newlines = @($DefaultHeaderLine1) + @($DefaultHeaderLine2) + $lines
        Set-Content -Path $file -Value $newlines
    }
    # Header is healthy, no action needed.
}


#####################################################################################
############################## Variables ############################################
#####################################################################################
[Bool]$EnabledRating = $true
[Bool]$KeepLocal = $true
$RegExCaseID = "^\w*[\s-]*\w*$"
$RegExEmpty = "^,{4,}$"
$profilepath = Get-ProfilePath
$filedir = "$profilepath\Documents"
$Current_Version = Get-LatestVersionId
$filename = "MyCaseTracker"
$filenosuffix = $filedir + "\" + $filename
$file = ("$filenosuffix" + ".csv")
$filexist = test-path $file


#####################################################################################
###################################### Main #########################################
#####################################################################################

if (!($false -ne $filexist)){
    New-Csv
}

Try{
    Start-Menu # -ErrorAction SilentlyContinue
}
Catch {
    Write-Warning "Oops,seems an error occurred, please check CSV file and download the latest version of script."
    $decision = Prompt-Confirm -action 'show error details?'
    if ($decision -eq 'y'){
        $ErrLine = $_.ScriptStackTrace
        $ErrType =  $($Error[0]).Exception.GetType().FullName
        $ErrDetail = $_.Exception.Message
        Write-Warning "`nLocation: `n$ErrLine`nType: `n$ErrType`nDetail:`n$ErrDetail"
    }
}



# Reference:
# http://woshub.com/read-write-excel-files-powershell/
# https://stackoverflow.com/questions/59402365/update-a-cell-in-a-excel-sheet-using-powershell