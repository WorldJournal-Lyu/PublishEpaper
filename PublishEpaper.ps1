<#

PublishEpaper.ps1

    2018-12-07 Initial Creation

#>

param([string]$Area = "TEST")

switch($Area){

    "BK" { 
        $pub = "BK"
        $newDate = 1
        $oldDate = -6
        break 
    }

    "LA" { 
        $pub = "LA"
        $newDate = 0
        $oldDate = -7
        break
    }

    "SF" {
        $pub = "SF"
        $newDate = 0
        $oldDate = -7
        break
    }

    "SE" {
        $pub = "SE"
        $newDate = 0
        $oldDate = -7
        break
    }

}

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Get-Module -ListAvailable WorldJournal.* | Remove-Module -Force
Get-Module -ListAvailable WorldJournal.* | Import-Module -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = "<"+(Get-WJEmail -Name jlee).MailAddress+">, <"+(Get-WJEmail -Name lyu).MailAddress+">"
$mailSbj    = $scriptName+"-"+$pub
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "-" + $pub + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

###################################################################################





# Define variables

$workDate = (Get-Date).AddDays($newDate)
$deleDate = (Get-Date).AddDays($oldDate)

$ftp      = Get-WJFTP -Name wjwebsite
[String]$ftpDir = Get-Content -Path ($scriptPath.Replace(".ps1",".FtpDir."+$pub+".txt"))
$pubName  = ($pub+"-"+$workDate.ToString("yyyy-MM-dd"))
$ftpWorkPath = $ftp.Path + $ftpDir + "/" + $workDate.ToString("yyyyMMdd") + "/"
$ftpDelePath = $ftp.Path + $ftpDir + "/" + $deleDate.ToString("yyyyMMdd") + "/"
$copyRight = "©" + (Get-Date).Year + " World Journal ALL RIGHTS RESERVED"
$pdfList = New-Object System.Collections.Generic.List[System.Object]
$jpgList = New-Object System.Collections.Generic.List[System.Object]
$bucketList = New-Object System.Collections.Generic.List[System.Object]
[String]$bucketName = Get-Content -Path ($scriptPath.Replace(".ps1",".Bucket.txt"))
Set-AWSCredential -ProfileName WJ_AWSProfile

Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal
Write-Log -Verb "deleDate" -Noun $deleDate -Path $log -Type Short -Status Normal
Write-Log -Verb "pubName" -Noun $pubName -Path $log -Type Short -Status Normal
Write-Log -Verb "ftpWorkPath" -Noun $ftpWorkPath -Path $log -Type Short -Status Normal
Write-Log -Verb "ftpDelePath" -Noun $ftpDelePath -Path $log -Type Short -Status Normal
Write-Log -Verb "bucketName" -Noun $bucketName -Path $log -Type Short -Status Normal

Write-Line -Length 50 -Path $log





# DELETE PDF
#   -- Input --
#   folder to delete on ftp ($ftpDelePath)
#   -- Output --
#   none

Write-Log -Verb "DELETE PDF" -Noun $ftpDelePath -Path $log -Type Long -Status System

$ftpTestPath = WebRequest-TestPath -Username $ftp.User -Password $ftp.Pass -RemoteFolderPath $ftpDelePath

if($ftpTestPath.Status -eq "Good"){

    $deleList = WebRequest-ListDirectory -Username $ftp.User -Password $ftp.Pass -RemoteFolderPath $ftpDelePath

    if($deleList.Status -eq "Good"){

        $deleList.list | ForEach-Object{

            $removeFile = (WebRequest-RemoveFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $ftpDelePath$_)

            if($removeFile.Status -eq "Good"){

                Write-Log -Verb $removeFile.Verb -Noun $removeFile.Noun -Path $log -Type Long -Status $removeFile.Status

            }elseif($removeFile.Status -eq "Bad"){

                $mailMsg = $mailMsg + (Write-Log -Verb $removeFile.Verb -Noun $removeFile.Noun -Path $log -Type Long -Status $removeFile.Status -Output String) + "`n"
                $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $removeFile.Exception -Path $log -Type Short -Status $removeFile.Status -Output String) + "`n"
                $hasError = $true

            }

        }

    }elseif($deleList.Status -eq "Bad"){

        $mailMsg = $mailMsg + (Write-Log -Verb $deleList.Verb -Noun $deleList.Noun -Path $log -Type Long -Status $deleList.Status -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $deleList.Exception -Path $log -Type Short -Status $deleList.Status -Output String) + "`n"
        $hasError = $true

    }

    $removeFolder = WebRequest-RemoveFolder -Username $ftp.User -Password $ftp.Pass -RemoteFolderPath $ftpDelePath

    if($removeFolder.Status -eq "Good"){

        Write-Log -Verb $removeFolder.Verb -Noun $removeFolder.Noun -Path $log -Type Long -Status $removeFolder.Status

    }elseif($removeFolder.Status -eq "Bad"){

        $mailMsg = $mailMsg + (Write-Log -Verb $removeFolder.Verb -Noun $removeFolder.Noun -Path $log -Type Long -Status $removeFolder.Status -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $removeFolder.Exception -Path $log -Type Short -Status $removeFolder.Status -Output String) + "`n"
        $hasError = $true

    }

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "DELETE PDF SKIPPED" -Noun $ftpDelePath -Path $log -Type Long -Status Warning -Output String) + "`n`n"

}

Write-Line -Length 50 -Path $log





# DOWNLOAD EXCEL
#   -- Input --
#   ftp user, password, remote file path, local file path
#   -- Output --
#   downloaded excel file path ($downloadExcel.Noun)

Write-Log -Verb "DOWNLOAD EXCEL" -Noun $downloadFrom -Path $log -Type Long -Status System

$downloadFrom = $ftpWorkPath + $pubName + ".xls"
$downloadTo   = $localTemp + $pubName + ".xls"
Write-Log -Verb "downloadFrom" -Noun $downloadFrom -Path $log -Type Short -Status Normal
Write-Log -Verb "downloadTo" -Noun $downloadTo -Path $log -Type Short -Status Normal

$downloadExcel = WebClient-DownloadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $downloadFrom -LocalFilePath $downloadTo

if($downloadExcel.Status -eq "Good"){

    Write-Log -Verb $downloadExcel.Verb -Noun $downloadExcel.Noun -Path $log -Type Long -Status $downloadExcel.Status

}elseif($downloadExcel.Status -eq "Bad"){

    $mailMsg = $mailMsg + (Write-Log -Verb $downloadExcel.Verb -Noun $downloadExcel.Noun -Path $log -Type Long -Status $downloadExcel.Status -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $downloadExcel.Exception -Path $log -Type Short -Status $downloadExcel.Status -Output String) + "`n"
    $hasError = $true

}

Write-Line -Length 50 -Path $log





# PARSE EXCEL
#   -- Input --
#   downloaded excel file path ($downloadExcel.Noun)
#   -- Output --
#   print date parsed from excel file name ($printDate)
#   publication data parsed from excel ($pubData)

Write-Log -Verb "PARSE EXCEL" -Noun $downloadExcel.Noun -Path $log -Type Long -Status System

Function ConvertExcelTo-Csv ($FullPath){

    $xls = $FullPath
    $csv = $FullPath.Replace("xls","csv")
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.DisplayAlerts = $False
    $wb = $Excel.Workbooks.Open($xls)
    foreach ($ws in $wb.Worksheets){ $ws.SaveAs($csv, 6) }
    $Excel.Quit()
    Write-Output $csv

}

if($downloadExcel.Status -eq "Good"){

    $basename  = ((Split-Path $downloadExcel.Noun -Leaf).Split("."))[0]
    $split     = $basename.Split("-")
    $printDate = Get-Date -Date (""+$split[1]+"-"+$split[2]+"-"+$split[3])
    $pubcode   = $split[0]

    try{

        $csvPath  = ConvertExcelTo-Csv -FullPath $downloadExcel.Noun
        $pubData  = Import-Csv -Path $csvPath -Encoding Default
        Write-Log -Verb "printDate" -Noun $printDate -Path $log -Type Short -Status Normal
        Write-Log -Verb "pubData.Count" -Noun (""+$pubData.Count+" lines") -Path $log -Type Short -Status Normal

    }catch{

        $mailMsg = $mailMsg + (Write-Log -Verb "PARSE EXCEL" -Noun $downloadExcel.Noun -Path $log -Type Long -Status Bad -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $upload.Exception -Path $log -Type Short -Status $upload.Status -Output String) + "`n"
        $hasError = $true

    }

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "PARSE EXCEL SKIPPED" -Noun $downloadExcel.Noun -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "Bad download excel status" -Path $log -Type Short -Status Bad -Output String) + "`n"
    $hasError = $true

}

Write-Line -Length 50 -Path $log





# DOWNLOAD PDF
#   -- Input --
#   publication data parsed from excel ($pubData)
#   print date parsed from excel file name ($printDate)
#   -- Output --
#   list of pdf successfully downloaded ($pdfList)

Write-Log -Verb "DOWNLOAD PDF" -Noun "pubData" -Path $log -Type Long -Status System

if($pubData.Count -gt 0){

    $pubData | ForEach-Object{

        $pdfName = $pubcode+$printDate.ToString("yyyyMMdd")+$_.Code+("00"+$_.Page).Substring(("00"+$_.Page).Length-2,2)
        $downloadFrom = $ftpWorkPath + $pdfName + ".pdf"
        $downloadTo   = $localTemp + $pdfName + ".pdf"
        Write-Log -Verb "downloadFrom" -Noun $downloadFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "downloadTo" -Noun $downloadTo -Path $log -Type Short -Status Normal

        $downloadPdf = WebClient-DownloadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $downloadFrom -LocalFilePath $downloadTo

        if($downloadPdf.Status -eq "Good"){

            Write-Log -Verb $downloadPdf.Verb -Noun $downloadPdf.Noun -Path $log -Type Long -Status $downloadPdf.Status
            $pdfList.Add($downloadPdf.Noun)

        }elseif($downloadPdf.Status -eq "Bad"){

            $mailMsg = $mailMsg + (Write-Log -Verb $downloadPdf.Verb -Noun $downloadPdf.Noun -Path $log -Type Long -Status $downloadPdf.Status -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $downloadPdf.Exception -Path $log -Type Short -Status $downloadPdf.Status -Output String) + "`n"
            $hasError = $true

        }

    }

    Write-Log -Verb "pdfList.Count" -Noun (""+$pdfList.Count+" files") -Path $log -Type Short -Status Normal

    if($pdfList.Count -eq $pubData.Count){

        Write-Log -Verb "DOWNLOAD PDF COUNT" -Noun "OK" -Path $log -Type Long -Status Good

    }else{

        $mailMsg = $mailMsg + (Write-Log -Verb "DOWNLOAD PDF COUNT" -Noun "NG" -Path $log -Type Long -Status Bad -Output String) + "`n"
        $hasError = $true

    }

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "DOWNLOAD PDF SKIPPED" -Noun "pubData" -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "Empty publication data" -Path $log -Type Short -Status Bad -Output String) + "`n"
    $hasError = $true

}

Write-Line -Length 50 -Path $log





# CONVERT PDF
#   -- Input --
#   list of pdf successfully downloaded ($pdfList)
#   -- Output --
#   list of jpg successfully converted ($jpgList)

Write-Log -Verb "CONVERT PDF" -Noun "pdfList" -Path $log -Type Long -Status System

if($pdfList.Count -gt 0){

    $pdfList | ForEach-Object{

        $convertFrom = $_
        $convertTo = $localTemp
        $convertJpg = $_.Replace(".pdf",".jpg")
        $convertJpgtn = $_.Replace(".pdf","-tn.jpg")

        Write-Log -Verb "convertFrom" -Noun $convertFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "convertTo" -Noun $convertTo -Path $log -Type Short -Status Normal

        magick mogrify -density 200 -format jpg -quality 90 -flatten -colorspace sRGB -gravity southwest -pointsize 12 -annotate +180+20 $copyRight -path $convertTo $convertFrom

        if(Test-Path $convertJpg){

            Write-Log -Verb "CONVERT" -Noun $convertJpg -Path $log -Type Long -Status Good
            $jpgList.Add($convertJpg)

        }else{

            Write-Log -Verb "CONVERT" -Noun $convertJpg -Path $log -Type Long -Status Bad

        }

        Copy-Item $convertJpg -Destination $convertJpgtn -ErrorAction Stop
        magick mogrify -resize 200x350 $convertJpgtn

        if(Test-Path $convertJpgtn){

            Write-Log -Verb "CONVERT" -Noun $convertJpgtn -Path $log -Type Long -Status Good
            $jpgList.Add($convertJpgtn)

        }else{

            Write-Log -Verb "CONVERT" -Noun $convertJpgtn -Path $log -Type Long -Status Bad

        }

    }

    Write-Log -Verb "jpgList.Count" -Noun (""+$jpgList.Count+" files") -Path $log -Type Short -Status Normal

    if($jpgList.Count -eq ($pdfList.Count*2)){

        Write-Log -Verb "CONVERT PDF COUNT" -Noun "OK" -Path $log -Type Long -Status Good

    }else{

        $mailMsg = $mailMsg + (Write-Log -Verb "CONVERT PDF COUNT" -Noun "NG" -Path $log -Type Long -Status Bad -Output String) + "`n"
        $hasError = $true

    }

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "CONVERT PDF SKIPPED" -Noun "pdfList" -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "Empty Pdf List" -Path $log -Type Short -Status Bad -Output String) + "`n"
    $hasError = $true

}

Write-Line -Length 50 -Path $log





# UPLOAD JPG
#   -- Input --
#   list of jpg successfully converted ($jpgList)
#   -- Output --
#   list of jpg successfully uploaded to bucket ($bucketList)


Write-Log -Verb "UPLOAD JPG" -Noun "jpgList" -Path $log -Type Long -Status System

if($jpgList.Count -gt 0){

    $jpgList | ForEach-Object{

        $basename = ((Split-Path $_ -Leaf).Split("."))[0]
        $area = $basename.Substring(0,2)
        $date = $basename.Substring(2,8)
        $jpgFile = ("/" + $area + "/" + $date + "/" + $baseName + ".jpg")
        $bucketFile = ("https://" + $bucketName + ".s3.amazonaws.com" + $jpgFile)
        Write-Log -Verb "jpgFile" -Noun $jpgFile -Path $log -Type Short -Status Normal
        Write-Log -Verb "bucketFile" -Noun $bucketFile -Path $log -Type Short -Status Normal

        try{

            Write-S3Object -BucketName $bucketName -Key $jpgFile -File $_ -CannedACLName public-read
            Write-Log -Verb "UPLOAD" -Noun $jpgFile -Path $log -Type Long -Status Good
            $bucketList.Add($bucketFile)

        }catch{

            Write-Log -Verb "UPLOAD" -Noun $jpgFile -Path $log -Type Long -Status Bad
            Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad

        }

    }

    Write-Log -Verb "bucketList.Count" -Noun (""+$bucketList.Count+" files") -Path $log -Type Short -Status Normal

    if($bucketList.Count -eq $jpgList.Count){

        Write-Log -Verb "UPLOAD JPG COUNT" -Noun "OK" -Path $log -Type Long -Status Good

    }else{

        $mailMsg = $mailMsg + (Write-Log -Verb "UPLOAD JPG COUNT" -Noun "NG" -Path $log -Type Long -Status Bad -Output String) + "`n"
        $hasError = $true

    }

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "UPLOAD JPG SKIPPED" -Noun "jpgList" -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "Empty Jpg List" -Path $log -Type Short -Status Bad -Output String) + "`n"
    $hasError = $true

}

Write-Line -Length 50 -Path $log





# CREATE JSON
#   -- Input --
#   downloaded excel file path ($downloadExcel.Noun)
#   -- Output --
#   json file path ($jsonPath)

Write-Log -Verb "CREATE JSON" -Noun $downloadExcel.Noun -Path $log -Type Long -Status System

if($downloadExcel.Status -eq "Good"){

    $jsonPath = Invoke-Expression -Command (($scriptPath.Replace(".ps1",".CreateJson.ps1"))+" -FilePath "+$downloadExcel.Noun)
    Write-Log -Verb "jsonPath" -Noun $jsonPath -Path $log -Type Short -Status Normal

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "CREATE JSON SKIPPED" -Noun "pubData" -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "Bad download excel status" -Path $log -Type Short -Status Bad -Output String) + "`n"
    $hasError = $true

}

Write-Line -Length 50 -Path $log





# UPLOAD JSON
#   -- Input --
#   json file path ($jsonPath)
#   -- Output --
#   SSH stream output to mail message

Write-Log -Verb "UPLOAD JSON" -Noun $jsonPath -Path $log -Type Long -Status System

if($jsonPath){

    if(Test-Path $jsonPath){

        $uploadResult = Invoke-Expression -Command (($scriptPath.Replace(".ps1",".UploadJson.ps1"))+" -FilePath "+$jsonPath+" -Pub "+$pub)

        if($hasError){
            $mailMsg = $mailMsg + "`n`n"
        }

        $mailMsg = $mailMsg + "SSH Result" + "`n"
        $mailMsg = $mailMsg + "---"
        $mailMsg = $mailMsg + $uploadResult

    }else{

        $mailMsg = $mailMsg + (Write-Log -Verb "UPLOAD JSON SKIPPED" -Noun $jsonPath -Path $log -Type Long -Status Bad -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "Json file doesn't exist" -Path $log -Type Short -Status Bad -Output String) + "`n"
        $hasError = $true

    }

}else{

    $mailMsg = $mailMsg + (Write-Log -Verb "UPLOAD JSON SKIPPED" -Noun $jsonPath -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Reason" -Noun "jsonPath variable doesn't exist" -Path $log -Type Short -Status Bad -Output String) + "`n"
    $hasError = $true

}





###################################################################################

Write-Line -Length 50 -Path $log

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $mailSbj }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam