param (
    [object]$WebhookData    
)

function Get-PdfText
{
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Path
    )

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

    try
    {
        $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $Path
    }
    catch
    {
        throw
    }

    $stringBuilder = New-Object System.Text.StringBuilder

    for ($page = 1; $page -le $reader.NumberOfPages; $page++)
    {
        $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page)
        $null = $stringBuilder.AppendLine($text) 
    }

    $reader.Close()

    return $stringBuilder.ToString()
}

function Export-MobilePhoneBillValues {
    param (
        [string]$Path
    )

    # Convert MobilePhone Bill to text
    $tmp = Get-PDFText -Path $Path
    # Get the dollar amount due for this bill
    $amount = [Regex]::Match($tmp, "(Total new charges in this bill \$)(\d+\.\d+)").captures.groups[2].Value

    # Get the date the bill is due
    [DateTime]$dateDue = [Regex]::Match($tmp, "(TOTAL DUE)\n(\d+ \w+ \d+)").Captures.Groups[2].Value
    [string]$dateDueUTC = $dateDue.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.ffffff+00:00")
    [string]$dateDue = $dateDue.ToUniversalTime().ToString("yyyy-MM-dd")

    # Get the billing period start and end dates
    $billingPeriod = [Regex]::Match($tmp, "(BILLING PERIOD)\n(\d+ \w+) \- (\d+ \w+)")

    [DateTime]$billingPeriodStart = $billingPeriod.Captures.Groups[2].Value
    [string]$billingPeriodStart = $billingPeriodStart.ToUniversalTime().ToString("yyyy-MM-dd")

    [DateTime]$billingPeriodEnd = $billingPeriod.Captures.Groups[3].Value
    [string]$billingPeriodEnd = $billingPeriodEnd.ToUniversalTime().ToString("yyyy-MM-dd")

    # Geneate the custom PSObject to hold our values.
    $obj = New-Object -TypeName PSObject
    $obj | Add-Member -MemberType NoteProperty -Name Provider -Value "Mobile Phone Provider"
    $obj | Add-Member -MemberType NoteProperty -Name Service -Value "Mobile Phone"
    $obj | Add-Member -MemberType NoteProperty -Name BillingPeriodStart -Value $billingPeriodStart
    $obj | Add-Member -MemberType NoteProperty -Name BillingPeriodEnd -Value $billingPeriodEnd
    $obj | Add-Member -MemberType NoteProperty -Name DateDueUTC -Value $dateDueUTC
    $obj | Add-Member -MemberType NoteProperty -Name DateDue -Value $dateDue
    $obj | Add-Member -MemberType NoteProperty -Name Amount -Value $amount
    $obj | Add-Member -MemberType NoteProperty -Name PaymentType -Value "Direct Debit"

    return $obj
}

# If runbook was called from a Webhook, $WebhookData will not be null
if ($WebhookData -ne $null) {

    # Collect properties of WebhookData
    #$WebhookName = $webhookData.WebhookName
    #$WebhookHeaders = $webhookData.RequestHeader
    $WebhookBody = $webhookData.RequestBody

    # Collect data - converted form JSON
    $bill = ConvertFrom-Json -InputObject $WebhookBody
    $billName = $bill.billName

    Write-Output "Processing JSON payload ..."
    Write-Output "Bill identified as $billName."
    
    # Define storage values
    $storageAccountName = 'mybillstorage' #LOWER CASE
    $storageAccessKey = '<Storage Key>'
    $ContainerName = 'attachments'
    $localFileDirectory = 'C:\'

    # Load the Azure Automation credential
    Write-Output "Logging to Azure using the automation account ..."
    $AzureOrgIdCredential = Get-AutomationPSCredential -Name 'myBillsCred'
    $Null = Add-AzureAccount -Credential $AzureOrgIdCredential  
    $Null = Select-AzureSubscription -SubscriptionName 'Pay-As-You-Go' 
    

    # Configure the Azure Blob Storage Context
    $ctx = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccessKey
   
    # Load the DLL 
    Write-Output "Loading itextsharp.dll"
    Add-Type -Path "C:\modules\user\itextsharp\itextsharp.dll"

    # Switch Bill Names for separate bill functions
    # RETURN VALUES: billName, billingPeriodStart, billingPeriodEnd, dateDueUTC, dateDue, amount, PaymentType
    Switch ($billName)
        {
            "MobilePhone" { 
                # Define the file name variables
                $BlobName = "$billName.pdf" 
                $localFile = $localFileDirectory + $BlobName 

                # Download the PDF from Azure Blob Storage
                $null = Get-AzureStorageBlobContent -Destination $localFile -Container $ContainerName -Blob $BlobName -Context $ctx

                # Call the export values function
                Write-Output "Executing: Export-MobilePhoneBillValues -Path $localfile"
                $billDetails = Export-MobilePhoneBillValues -Path $localFile
            }
            "Gas" {
                # Gas bill parsing code here ...
            }
            "Electric" { 
                # Electricity bill parsing code here ...
            }
        }
    
    # Convert the function output to JSON
    $json = $billDetails | ConvertTo-json

    # Execute the 'Take Action' Flow by calling it's webhook.
    $uri = '<Webhook URL>'
    Write-Output "Executing Web Request to 'Take Action' Flow"
    Invoke-RestMethod -Uri $uri `
                        -Method Post `
                        -Body $json `
                        -ContentType 'application/json'

    # Clean up 
    # Remove the PDF from Azure Blob Storage
    $null = Remove-AzureStorageBlob -Container $ContainerName -Blob $BlobName -Context $ctx

    # Remove the PDF from the local system
    Remove-Item $localFile

} else {
    Write-Output "ERROR: Runbook must be started from a webook."
}
