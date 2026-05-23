
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Body,
    [Parameter(Mandatory = $true)]
    [string] $Subject,
    [Parameter(Mandatory = $true)]
    [string] $FromEmailAddress,
    [Parameter(Mandatory = $true)]
    [string] $ToEmailAddresses,
    [Parameter(Mandatory = $true)]
    [string] $SmtpServer,
    [Parameter(Mandatory = $true)]
    [string] $ReleaseArtifactName,
    [Parameter(Mandatory = $true)]
    [string] $TXTBuildDropPath
)

Function SendEmail() {
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $ToEmailAddresses,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Subject,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Body,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $SmtpServer,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Attachment
    )

    $Message = new-object System.Net.Mail.MailMessage 
    $Message.From = $FromEmailAddress 
    $Message.To.Add($ToEmailAddresses)
    $Message.IsBodyHtml = $True 
    $Message.Subject = $Subject 
    $attach = new-object System.Net.Mail.Attachment($Attachment)

    Write-Host "Adding Attachment"
    $Message.Attachments.Add($attach) 
    $Message.body = $Body
    $Smtp = new-object System.Net.Mail.SmtpClient($SmtpServer) 

    Write-Host "Sending the invoice error report to the e-mail(s): " $ToEmailAddresses
    $Smtp.Send($Message)

    Write-Host "##[section] E-mail sent!"
}

$now = Get-Date

if (($now -ge $now.Date) -and ($now -le $now.Date.AddHours(8))) {
    $CSVInvoiceErrorListFilePath = $TXTBuildDropPath + "\" + "Invoice_ErrorList.csv"

    Write-Host "Checking for the invoice error reporting file"
    if (Test-Path $CSVInvoiceErrorListFilePath -PathType Leaf) {   
        Write-Host "Found the file: Invoice_ErrorList.csv"
        SendEmail -ToEmailAddresses $ToEmailAddresses -Subject $Subject -Body $Body -SmtpServer $SmtpServer -Attachment $CSVInvoiceErrorListFilePath
    }
    else {
        Write-Host "##[section] Hurray! No invoices with error found!"
    }
}
else {
    Write-Host "The email will be sent only between 00:00 and 08:00"
}