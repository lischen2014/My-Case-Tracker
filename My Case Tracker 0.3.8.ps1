Clear-Host
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

#####################################################################################
############################# Update History ########################################
#####################################################################################

$UpdateHistory =@"

======================================= Start =======================================

    v0.1.0 = Initial version.
    v0.1.1 = Fixed a bug in version display.
    v0.1.2 = Add feature to delete last record.
	         Separated add case tip.
             Added test option to bulk import data.
    v0.1.3 = Modified prompt message.
    v0.1.4 = Added time before added message.
    v0.1.5 = Shows case note in added message
    v0.1.6 = Remove blank lines from csv when view data.
             Added monthly work history
             Renamed 'TodayWork' to 'DailyWork'.
    v0.1.7 = Fixed a bug in time record.
             Added case detail in daily/monthly work review.
    v0.1.8 = Modified prompt code.
    v0.1.9 = Auto detect OneDrive linked profile.
             Selected current month as default month in Monthly history.
    v0.2.0 = Modified Update History code, enhanced performance.
    v0.2.1 = Add new feature of fix csv header.
             Modified some text.
    v0.2.2 = Fixed a bug in remove specific case.
             Fixed a bug in check default path.
             Modified some text.
             Hide test option from main menu.
    v0.2.3 = Modified add new case logic.
    v0.2.4 = Fixed a bug of daily history.
             Fixed a bug of date in monthly history.
    v0.2.5 = Changed score calculation, lower chat & phone to 1.
             Fixed a bug in Daily & monthly history.
    V0.2.6 = Fixed a bug in remove empty line.
    v0.2.7 = Add new feature of fix header missing.
    v0.2.8 = Fixed bug in add new case.
    v0.2.9 = Fixed a bug in remove empty line.
    v0.3.0 = Add a feature of debug.
             Trim input message.
    v0.3.1 = Removed prompt in View-AllWork.
    v0.3.2 = Add a feature of case details with GUI.
    v0.3.3 = Add a feature of recognize phone as keyword.
    v0.3.4 = Add support of '-' in case id.
    v0.3.5 = Add a feature of auto remove line break.
    v0.3.6 = Fixed a bug of wrong monthly display.
             Changed view to percentage, default 35 for 100%.
             Changed monthly view to average score.
    v0.3.7 = Text modification.
             Fixed a bug in Monthly view.
             Changed version ID, now started with 0.1.
             Modify version check function.
    v0.3.8 = Text modification.

           

    Latest Releases: 
    Releases: https://github.com/lischen2014/My-Case-Tracker

                                                   Author: Leon
                                                   Email: leon2014@vip.qq.com
======================================== End ========================================

"@




#####################################################################################
################################ Functions ##########################################
#####################################################################################

function Show-Menu{
    param (
    [string]$Title = 'Menu'
    )
    Write-Host ""
    Write-Host "=============== $Title ==============="
    write-host ""
    Write-Host "1: Press '1' to add a case."
    Write-Host "2: Press '2' to remove a case."
    Write-Host "3: Press '3' to view daily history."
    Write-Host "4: Press '4' to view monthly history."
    Write-Host "5: Press '5' to view all history."
    # Write-Host "7: Press '7' to Import test data.(Test option, do not use!)"
    Write-Host "8: Press '8' to open CSV folder from file explorer."
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
        Write-Host ""
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
                View-DailyWork
            }
            '4'{
                view-MonthlyWork
            }
            '5'{
                View-AllWork
            }
            
            # space for test option

            '8'{
                # open csv file location
                explorer $Filedir
            }
            '9'{
                $UpdateHistory
                pause
            }
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
        # Modify custom path here with $keeplocal set to $true.
        $profilepath = "C:\Users\$env:USERNAME"
    }

    return $profilepath
}


function Get-LatestVersionId {
    $pattern = 'v(\d+(\.\d+){0,3})'
    $matches = [regex]::Matches($UpdateHistory, $pattern)
    
    if ($matches.Count -gt 0) {
        $latestMatch = $matches[$matches.Count - 1]
        return $latestMatch.Groups[0].Value # 0 with v, 1 without v.
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
    $NewSheet.cells.item(2,4) = 'Type'
    $NewSheet.cells.item(2,5) = 'Note'
    # Save the file
    try{
        $NewWorkbook.SaveAs("$filePath",[Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV) # xlCSV specifies the CSV file format
        write-host "CSV is created, the path is:"
        write-host $File -ForegroundColor Cyan
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

    # Add Type
    if ($case.case -match "Phone"){
        $case.type = "Phone"
    }
    elseif (($case.case -match "INC") -or ($case.case -match "REQ") -or ($case.case -match "RITM")){
        $case.type = "Written"
    }
    elseif($case.case -match "IMS"){
        $case.type = "Chat"
    }
    else{
        $case.type = ((Read-Host "Please type case type(use written, phone or chat, default is written.)").trim()) -replace "\r?\n", " "
    }

    # Add Note
    $case.note = ((Read-Host "[Optional] Please type case additional information").trim()) -replace "\r?\n", " "
    
    if(!$case.type){
        $case.type = "Written"
    }

    $case.content = $case.date, $case.time, $case.case, $case.type, $case.note -join ','
    write-host ""
    try{
        Transmit-Case -case $case
        if(!$case.note){
            Write-HostWithTime -Message "Message: $($case.case) is recorded." -ForegroundColor Cyan
        }
        else{
            Write-HostWithTime -Message "Message: $($case.case) - $($case.note) is recorded." -ForegroundColor Cyan
        }
        
    }
    catch{$e}
}


function Transmit-Case{
    param(
        $case
    )
    foreach ($singlecase in $case.content){
        $singlecase | add-content -path $File
    }
}


function View-DailyWork{
    Remove-EmptyLine
    $csv = Import-Csv $File | Select-Object *

    # Parse data
    $TodayWork = $csv | Where-Object {([DateTime]$_.Date) -eq $date} 
    $Today = [ordered]@{}

    # Calculate how many written case received
    $Today.WrittenDetail = $TodayWork | Where-Object {$_.type -match "Case|Written|NowIT"}
    $Today.Written = ($Today.WrittenDetail | Measure-Object).Count

    # Calculate how many chat case received
    $Today.ChatDetail =$TodayWork | Where-Object {$_.Type -eq "Chat"}
    $Today.Chat = ($Today.ChatDetail | Measure-Object).Count

    # Calculate how many phone case received
    $Today.PhoneDetail = $TodayWork | Where-Object {$_.Type -eq "Phone"}
    $Today.Phone = ($Today.PhoneDetail | Measure-Object).Count
    # sum, calculate the grade
    $Today.Score = ‘{0:p1}’ -f (($Today.Written + $Today.Chat + $Today.Phone)/$target)

    # Display the current data
    $TodayReview = New-Object PSObject -property @{
        Date = $date
        Score = $today.Score
        Written = $today.Written
        Chat = $today.Chat
        Phone = $today.Phone
    }

    # Summary
    $TodayReview | Format-Table -Property Date, Score, Written, Chat, Phone -AutoSize -Wrap
    Write-host "How Score counts:  Score = (Written + Chat + Phone)/Daily Target"
    Display-Details -table $TodayWork
}


function View-MonthlyWork{
    Remove-EmptyLine

    # Select time range
    try {
        $SOMUserInput = (Read-Host "Enter a month,current month by default: (Eg: 2023-01)").trim()
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
    
    $csv = Import-Csv $File | Where-Object -FilterScript {([DateTime]::Parse($_."Date") -ge $StartOfMonth) -and ([DateTime]::Parse($_."Date") -lt $EndOfMonth)} 

    # Parse data
    $MonthWork = $csv
    $Month = [ordered]@{}
    # Calculate how many written case received
    $Month.WrittenDetail = $MonthWork | Where-Object {$_.type -match "Case|Written|NowIT"}
    $Month.Written = ($Month.WrittenDetail | Measure-Object).Count

    # Calculate how many chat case received
    $Month.ChatDetail =$MonthWork | Where-Object {$_.Type -eq "Chat"}
    $Month.Chat = ($Month.ChatDetail | Measure-Object).Count

    # Calculate how many phone case received
    $Month.PhoneDetail = $MonthWork | Where-Object {$_.Type -eq "Phone"}
    $Month.Phone = ($Month.PhoneDetail | Measure-Object).Count
    # sum, calculate the grade
    # $Month.Score = ‘{0:p1}’ -f (($Month.Written + $Month.Chat*1 + $Month.Phone*1)/35)
    
    # Calculate Average Score
    $Month.Days = ($Monthwork.Date | Get-Unique).count
    $Month.AvgScore = ‘{0:p1}’ -f (($Month.Written + $Month.Chat*1 + $Month.Phone*1)/$target/$Month.Days)

    # Display the current data
    $MonthReview = New-Object PSObject -property @{
        Date = ([String]$StartOfMonth).Substring(0,3)+([String]$StartOfMonth).Substring(6,4)
        Days = $Month.Days
        AvgScore = $Month.AvgScore
        Written = $Month.Written
        Chat = $Month.Chat
        Phone = $Month.Phone
    }

    # Summary
    $MonthReview | Format-Table -Property Date, Days, AvgScore, Written, Chat, Phone -AutoSize -Wrap 
    Write-host "How Score counts:  Score = (Written + Chat + Phone)/Daily Target/Record Days"
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
    $decision = Prompt-Confirm -action 'display work details'
    if ($decision -eq 'y'){
        # $table | Format-Table -AutoSize
        $table | Out-GridView
    }
    else{
        write-host "Cancelled,back to main menu."
    }
}


function Remove-LastLine{
    $lines = get-content -path $File
    $lines = $lines[0..($lines.count-2)]
    Set-Content $File -Value $lines
}


function Remove-EmptyLine{
    # Remove csv empty lines and txt empty lines
    (gc $file) | ? {($_.trim() -notmatch $RegExEmpty) -and ($_.trim() -ne "") } | Set-Content $file
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

[Bool]$EnabledRating = $false
[Bool]$KeepLocal = $false
$target = 35
$RegExCaseID = "^\w*[\s-]*\w*$"
$RegExEmpty = "^,{4,}$"
$profilepath = Get-ProfilePath
$Filedir = "$profilepath\Documents"
$Current_Version = Get-LatestVersionId
$Filename = "MyCaseTracker"
$filePath = $Filedir + "\" + $Filename
$File = ("$filePath" + ".csv")
$Filexist = test-path $File



#####################################################################################
###################################### Main #########################################
#####################################################################################

if (!($false -ne $Filexist)){
    New-Csv
}

Try{
    Start-Menu # -ErrorAction SilentlyContinue
}
Catch {
    write-host ""
    Write-Warning "Oops,seems an error occurred, please check CSV file and download the latest version of script.`n"
    write-host "You can download latest version from here: `nhttps://git.build.ingka.ikea.com/LEJIA3/My-Case-Tracker`n"
    $decision = Prompt-Confirm -action 'show error details?'
    if ($decision -eq 'y'){
        $ErrLine = $_.ScriptStackTrace
        $ErrType =  $($Error[0]).Exception.GetType().FullName
        $ErrDetail = $_.Exception.Message
        Write-Warning "`nLocation: `n$ErrLine`nType: `n$ErrType`nDetail:`n$ErrDetail"
    }
}


