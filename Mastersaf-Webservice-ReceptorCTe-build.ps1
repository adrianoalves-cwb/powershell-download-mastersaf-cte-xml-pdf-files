
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [String]$Company,
    [Parameter(Mandatory = $true)]
    [String]$UserName,
    [Parameter(Mandatory = $true)]
    [String]$Userkey,
    [Parameter(Mandatory = $true)]
    [String]$MastersafWebServiceURL,
    [Parameter(Mandatory = $true)]
    [String]$DaysInThePast
)

$TXTSavingPath = $env:BUILD_ARTIFACTSTAGINGDIRECTORY

Switch ($Company) {
    "NOVA" {
        $CompanyCNPJ = "78150693422047"
        $CompanyIE = "6248091735";
        break
    }
    "ZILO" {
        $CompanyCNPJ = "50287416933158"
        $CompanyIE = "9157362048";
        break
    }
    "TREX" {
        $CompanyCNPJ = "36421980574763"
        $CompanyIE = "847120536904";
        break
    }
    "LUMA" {
        $CompanyCNPJ = "91643578210859"
        $CompanyIE = "3026915847";
        break
    }
}

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

    $Request = Invoke-WebRequest -UseBasicParsing -Method POST -Uri $LoginUrl -Headers $LoginHeaders  -Body $Json #-Verbose 

    $Result = $Request | ConvertFrom-Json

    if (!($Result.accessToken)) {
        Write-Host "##[error] Failed to Retrieve Token."
    }

    return $Result.accessToken
}

Function IsTodayLastDayOfMonth {
    $Today = Get-Date
    $FirstDayOfMonth = Get-Date $Today -Day 1
    $LastDayOfMonth = Get-Date $FirstDayOfMonth.AddMonths(1).AddDays(-1)

    if ($Today -eq $LastDayOfMonth) {
        return $true
    }

    return $false
}

Function QueryReceptorDocumentWebService {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $MastersafWebServiceURL,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Token,
        [Parameter(Mandatory = $true, Position = 2)]
        [string] $CompanyCNPJ,
        [Parameter(Mandatory = $true, Position = 3)]
        [string] $CompanyIE,
        [Parameter(Mandatory = $true, Position = 3)]
        [string] $OffSet
    )

    $AuthorizationHeaders = @{
        Authorization = "Bearer $Token"
    }

    $Date = Get-Date((get-date ).AddDays($DaysInThePast)) -Format "dd/MM/yyyy"

    if (IsTodayLastDayOfMonth) {
        $Date = Get-Date((get-date ).AddDays(-32)) -Format "dd/MM/yyyy"
    }

    $Yesterday = Get-Date((get-date ).AddDays(-1)) -Format "dd/MM/yyyy"

    $PeriodoInicial = $Date + " 00:00"
    $PeriodoFinal = $Yesterday + " 23:59"

    Write-Host "##[command]Querying the WebService for date range: "$PeriodoInicial "and" $PeriodoFinal

    $Url = $MastersafWebServiceURL + "/api/recebimento/getListagem?cnpj=" + $CompanyCNPJ + "&ie=" + $CompanyIE + "&periodoInicial=" + $PeriodoInicial + "&periodoFinal=" + $PeriodoFinal + "&offset=" + $OffSet + "&maxResults=100"

    Write-Host "##[command]WebService Query:" $Url

    try {
        $Request = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $Url -Headers $AuthorizationHeaders -ContentType "application/json" -Verbose
        $Result = $Request | ConvertFrom-Json
        return $Result
    }
    catch {
        Write-Host "An Error has ocurred when querying the Mastersaf WebService! Waiting 10 minutes to retry..."

        return $Result
    }
}

# 2 - Query WebService
Function GetListOfInvoices {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $MastersafWebServiceURL,
        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Token,
        [Parameter(Mandatory = $true, Position = 2)]
        [string] $CompanyCNPJ,
        [Parameter(Mandatory = $true, Position = 3)]
        [string] $CompanyIE
    )

    $OffSet = 0

    $Result = QueryReceptorDocumentWebService -MastersafWebServiceURL $MastersafWebServiceURL -Token $Token -CompanyCNPJ $CompanyCNPJ -CompanyIE $CompanyIE -OffSet $OffSet

    $CounterResult = 0
    $CounterResult = $Result.count

    Write-Host "Number of invoices found: " $Result.count

    if (($CounterResult -eq 0) -or ($CounterResult -le 100)) {
        Write-Host "Found " $Result.count " invoices from the Mastersaf WebService."
        return $Result.result
    }
    
    $ResultList = $Result.result

    do {
        $OffSet = $OffSet + 100

        $Result = QueryReceptorDocumentWebService -MastersafWebServiceURL $MastersafWebServiceURL -Token $Token -CompanyCNPJ $CompanyCNPJ -CompanyIE $CompanyIE -OffSet $OffSet

        $ResultList += $Result.result

        $CounterResult = $CounterResult - 100

    } While ($CounterResult -ge 100)

    Write-Host "Found " $ResultList.count " invoices from the Mastersaf WebService."

    return $ResultList
}

Function CreateTXTFile($InvoiceKeys) {
    $TXTSavingPathWithName = $TXTSavingPath + "\" + $CompanyCNPJ + ".txt"
    New-Item $TXTSavingPathWithName
    Set-Content $TXTSavingPathWithName $InvoiceKeys
}

Function CreateCSVErrorFile($InvoiceKeys) {
    $TXTSavingPathWithName = $TXTSavingPath + "\" + "Invoice_ErrorList.csv"

    if (!(Test-Path $TXTSavingPathWithName -PathType Leaf)) {
        New-Item $TXTSavingPathWithName
        Add-Content $TXTSavingPathWithName "SEP=;"
        Add-Content $TXTSavingPathWithName "InvoiceDateTime; Company; InvoiceKey; InvoiceStatus"
    }
    
    Add-Content $TXTSavingPathWithName $InvoiceKeys
}

Function GetInvoiceKey($InvoiceKeyWithPrefix) {
    $InvoiceKey = $InvoiceKeyWithPrefix.Replace('CTe', '')

    return $InvoiceKey
}

$global:Token = GetToken($MastersafWebServiceURL)

if ($Token) {
    $counter = 0
    $ErrorCounter = 0

    $InvoiceKeys = @()
    $ErrorStatusInvoices = @()

    $ListOfInvoices = GetListOfInvoices -MastersafWebServiceURL $MastersafWebServiceURL -Token $Token -CompanyCNPJ $CompanyCNPJ -CompanyIE $CompanyIE

    Foreach ($invoice in $ListOfInvoices) {
        $counter++

        $InvoiceStatus = $invoice.situacao

        if (($InvoiceStatus -eq "Autorizado") -or ($InvoiceStatus -eq "Cancelado")) {
            $InvoiceKey = GetInvoiceKey($invoice.chaveAcesso)
            $InvoiceKeys += $InvoiceKey
                 
            Write-Host $counter "-" $InvoiceKey " - Invoice Status:" $InvoiceStatus 
        
        }
        else {
            $ErrorCounter++
            
            $InvoiceKey = GetInvoiceKey($invoice.chaveAcesso)
            
            $TXTErrorText = $invoice.dataEmissao + ";" + "=""${CompanyCNPJ}""" + ";" + "=""${InvoiceKey}""" + ";" + $invoice.situacao
            $ErrorStatusInvoices += $TXTErrorText

            Write-Host "##[error]" $counter "-" $InvoiceKey " - Invoice Status:" $InvoiceStatus 
        }
    }

    if ($counter -gt 0) {
        CreateTXTFile($InvoiceKeys)
    }

    if ($ErrorCounter -gt 0) {
        CreateCSVErrorFile($ErrorStatusInvoices)
    }

}