
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [String]$Company,
    [Parameter(Mandatory = $true)]
    [String]$UserName,
    [Parameter(Mandatory = $true)]
    [String]$Userkey,
    [Parameter(Mandatory = $true)]
    [String]$CompanyCNPJ,
    [Parameter(Mandatory = $true)]
    [String]$CompanyIE,
    [Parameter(Mandatory = $true)]
    [String]$SavingFolderName,
    [Parameter(Mandatory = $true)]
    [String]$ApplicationName,
    [Parameter(Mandatory = $true)]
    [String]$MastersafWebServiceURL,
    [Parameter(Mandatory = $true)]
    [String]$XMLSavingFolderPath,
    [Parameter(Mandatory = $true)]
    [String]$PDFSavingFolderPath,
    [Parameter(Mandatory = $true)]
    [String]$LogFolderPath,
    [Parameter(Mandatory = $true)]
    [String]$HistoryFileFolderPath,
    [Parameter(Mandatory = $true)]
    [string] $FileType,
    [Parameter(Mandatory = $false)]
    [string] $TXTBuildDropPath,
    [Parameter(Mandatory = $false)]
    [bool] $EnableSapLes = $false
)

Write-Host "Company: " $Company
Write-Host "CompanyCNPJ : " $CompanyCNPJ

#Login
Function GetToken($MastersafWebServiceURL) {
    $LoginUrl = $MastersafWebServiceURL + "/api/login"

    $Body = @{
        nomeUsuario = $UserName;
        chave       = $Userkey;
    }

    $Json = $Body | Convertto-JSON

    $LoginHeaders = @{
        "Content-Type" = "application/json"
        "Accept"       = "application/json;odata=fullmetadata"
    }

    $Request = Invoke-WebRequest -UseBasicParsing -Method POST -Uri $LoginUrl -Headers $LoginHeaders  -Body $Json -Verbose 

    $Result = $Request | ConvertFrom-Json

    if (!($Result.accessToken)) {
        WriteLog -Level "ERROR" -Message "Failed to Retrieve Token."
    }

    return $Result.accessToken
}

#Log function
Function WriteLog {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Level,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Message
    )

    $private:Level = "[" + $Level + "]"
    $LogNameDateFormat = (Get-Date).ToString('yyyyMMdd')

    $TimeStamp = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss tt')

    $LogFile = $LogFolderPath + "\" + $ApplicationName + "-" + $Company + "-" + $LogNameDateFormat + ".txt"

    $Line = $TimeStamp + "- " + $Level + " - " + $Message

    if (Test-Path $LogFile) {
        Add-Content $LogFile -Value $Line
    }
    Else {
        $Line | Out-File -FilePath $LogFile
    }
}

Function CheckHistoryFile {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $InvoiceKey,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $HistoryFileNamePath
    )

    if (Test-Path $HistoryFileNamePath -PathType Leaf) {
        foreach ($line in Get-Content $HistoryFileNamePath) {
            if ($line -match $InvoiceKey) {
                return $true
            }
        }
    }

    return $false
}

Function AddToHistoryFile {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $FileName,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $HistoryFileNamePath,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $HistoryFilePath
    )

    $TimeStamp = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss tt')
    $Line = $TimeStamp + " - " + $FileName

    if (Test-Path $HistoryFilePath) {
        Add-Content $HistoryFileNamePath -Value $Line
    }
    Else {
        $Line | Out-File -FilePath $HistoryFileNamePath
    }
}

Function Archive30DaysHistoryFiles {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $HistoryFilePath
    )

    $HistoryFiles = Get-ChildItem -Path $HistoryFilePath\*.txt
    Foreach ($file in $HistoryFiles) {
        $HistoryFileToCheck = $CompanyCNPJ + ".txt"

        If ($file.Name -eq $HistoryFileToCheck) {
            $FileCreationDate = $file.CreationTime.ToString('yyyyMMdd')
            $Past30Days = Get-Date((get-date ).AddDays(-30)) -Format "yyyyMMdd"

            If ($file.CreationTime.ToString('yyyyMMdd') -le $Past30Days) {
                Write-Host "Archiving the History Files"

                $OldFileName = $file.FullName
                $NewFileName = $HistoryFilePath + "\" + $CompanyCNPJ + "_" + $FileCreationDate + ".txt"

                Write-Host "OldFileName: " $OldFileName
                Write-Host "NewFileName: " $NewFileName

                #Rename-Item -Path $file.FullName -NewName $NewFileName
                Rename-Item -Path $OldFileName -NewName $NewFileName
            }
        }
    }
}

Function DownloadFile {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Token,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $InvoiceKey,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $FileExtension,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Uri,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $FileNameWithPath
    )

    $AuthorizationHeaders = @{
        Authorization = "Bearer $Token"
    }

    if (Test-Path $FileNameWithPath -PathType Leaf) {
        Remove-Item $FileNameWithPath
    }

    Write-Host "Bearer: $Token"
    Write-Host "Donwloading File Command: Invoke-WebRequest -Method GET -Uri $Uri -Headers $AuthorizationHeaders -OutFile $FileNameWithPath"

    Invoke-WebRequest -UseBasicParsing -Method GET -Uri $Uri -Headers $AuthorizationHeaders -OutFile $FileNameWithPath

}

$global:Token = GetToken($MastersafWebServiceURL)

if ($Token) {
    $counter = 1
    $ArtifactFileExists = $false
    $ArtifactFileName = $CompanyCNPJ + ".txt"

    Foreach ($artifacFile in Get-ChildItem $TXTBuildDropPath) {
        if ($artifacFile -match $ArtifactFileName) {
            Write-Host "Artifact TXT File found: "$artifacFile
            $ArtifactFileExists = $true
            $TXTFilesPath = $TXTBuildDropPath + "\" + $artifacFile
            break
        }
    }

    if ($ArtifactFileExists) {
        $DownloadCounter = 0

        Foreach ($invoiceKey in Get-Content $TXTFilesPath) {
            if ($FileType -eq "XML") {
                $FileExtension = ".xml"
                $Uri = $MastersafWebServiceURL + "/api/recebimento/" + $InvoiceKey + "/getXml"
                $FileName = $InvoiceKey + "-document.xml"
                $FileNameWithPath = $XMLSavingFolderPath + "\" + $SavingFolderName + "\" + $FileName
                $HistoryFilePath = $HistoryFileFolderPath + "\xml"
                $HistoryFileNamePath = $HistoryFileFolderPath + "\xml\" + $CompanyCNPJ + ".txt"

            }
            else {
                $FileExtension = ".pdf"
                $Uri = $MastersafWebServiceURL + "/api/recebimento/" + $InvoiceKey + "/getDacte"
                $FileName = $InvoiceKey + "-document.pdf"
                $FileNameWithPath = $PDFSavingFolderPath + "\" + $SavingFolderName + "\" + $FileName
                $HistoryFilePath = $HistoryFileFolderPath + "\pdf"
                $HistoryFileNamePath = $HistoryFileFolderPath + "\pdf\" + $CompanyCNPJ + ".txt"
            }

            $InvoiceKeyInHistoryFile = CheckHistoryFile -InvoiceKey $invoiceKey -HistoryFileName $HistoryFileNamePath

            if ($InvoiceKeyInHistoryFile -eq $false) {
                
                DownloadFile -Token $Token -InvoiceKey $InvoiceKey -FileExtension $FileExtension -Uri $Uri -FileNameWithPath $FileNameWithPath
                $DownloadCounter ++

                Write-Host $counter " - " $DownloadCounter " Downloaded and Saved the File: "$FileName

                AddToHistoryFile -FileName $FileName -HistoryFileName $HistoryFileNamePath -HistoryFilePath $HistoryFilePath

            }
            else {
                Write-Host $counter "##[section]File already in the History report: "$FileName
            }        

            $counter++

        }

        Write-Host "Number of Downloaded Files: " $DownloadCounter

        #Archiving the History Files if they are older than 30 days
        Archive30DaysHistoryFiles -HistoryFilePath $HistoryFilePath
    }
    else {
        Write-Host "##[section]Artifact TXT File could NOT be Found or had no data for the day."
    }
}