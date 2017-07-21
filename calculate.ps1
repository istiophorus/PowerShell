

$rcbItemsMap = @{ "RCFL3OIL" = "https://www.rcb.at/produkt/factor/?ID_NOTATION=68555282&ISIN=AT0000A0WJH0"; "RCCOPPOPEN" = "https://www.rcb.at/produkt/participation/?ID_NOTATION=170707925&ISIN=AT0000A1NR81"; "RCWHTAOPEN" = "https://www.rcb.at/produkt/participation/?ID_NOTATION=62781084&ISIN=AT0000A05VT0" }

function Get-RcbItemData
{
    Param (
        [parameter(Mandatory = $true)][string] $itemName
    )    

    Write-Host "[Get-RcbItemData] $itemName"

    $url = $rcbItemsMap[$itemName]

    Write-Host "[Get-RcbItemData] Loading data from $url"

    $data = Invoke-WebRequest $url -Method GET

    [array] $items = ([regex]::matches($data, "valueFilter:priceFilter2"">(.+)</span><br/>") | %{$_.value})

    $prefix = "valueFilter:priceFilter2"">"
    $sufix = "</span><br/>"

    if ($null -eq $items) {
        $prefix = "valueFilter:priceFilter3"">"
        $sufix = "</span><br/>"
        $items = ([regex]::matches($data, "valueFilter:priceFilter3"">(.+)</span><br/>") | %{$_.value})
    }

    Write-Host $items

    if ($null -eq $items) {
        return -1
    }

    if ($items.Length -lt 1) {
        return -1
    }

    $lastItemText = [string]$items[0]

    Write-Host "Matching $lastItemText"

    $extracted = $lastItemText.Substring($prefix.Length, $lastItemText.Length - $prefix.Length - $sufix.Length)

    Write-Host "Extracted $extracted"

    $priceValue = [System.Decimal]::Parse($extracted)

    Write-Host "Parsed $priceValue"

    $priceValue
}

function Get-IngItemData
{
    Param (
        [parameter(Mandatory = $true)][string] $itemName
    )

    Write-Host "[Get-IngItemData] $itemName"

    $newItemName = "PLINGNV" + $itemName.Substring(7, $itemName.Length - 7)

    $baseName = $itemName.Substring(4, 3)

    Write-Host "[Get-IngItemData] New item name $baseName $newItemName"

    $url = "https://www.ingturbo.pl/services/product/$newItemName/chart?period=intraday"

    Write-Host "[Get-IngItemData] Loading data from $url"

    $data = Invoke-WebRequest $url -Method GET

    $receivedData = ConvertFrom-Json -InputObject $data

    $bidValue = $receivedData.BidQuotes[$receivedData.BidQuotes.Length - 1][1]

    $askValue = $receivedData.AskQuotes[$receivedData.AskQuotes.Length - 1][1]

    $refValue = $receivedData.ReferenceQuotes[$receivedData.ReferenceQuotes.Length - 1][1]

    Write-Host "[Get-IngItemData] $itemName $bidValue $askValue $refValue"

    $result = ($bidValue, $askValue, $refValue)

    $result
}

 #(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

$data = Import-csv -Path "..\export.csv" -Delimiter ';'

$cashOut = 0.0
$cashIn = 0.0
$commission = 0.0
$commissionRate = 0.0039
$balance = 0.0
$expectedIncomeRate = 0.10

foreach ($item in $data)
{
    $itemName = $item."Papier"

    Write-Host "Item $itemName"

    $currentPrice = 0

    if ($itemName.StartsWith("INTL"))
    {
        $resultValues = Get-IngItemData $itemName

        $currentPrice = $resultValues[0]
    }
    elseif ($itemName.StartsWith("RC"))
    {
        $currentPrice = Get-RcbItemData $itemName
    }
    else {
        continue
    }

    if ($currentPrice -le 0) {
        continue
    }

    $itemsCount = [int]$item."Pozycja"

    $currentOut = ([double]$item."Kurs K") * ($itemsCount)

    $currentCommissionOut = $currentOut * $commissionRate

    $currentIn = [System.Math]::Round($currentPrice * $itemsCount, 2)

    $currentCommissionIn = [System.Math]::Round($currentIn * $commissionRate, 2)

    $currentOutAll = [System.Math]::Round($currentOut + $currentCommissionOut + $currentCommissionIn, 2)

    $currentBalance = $currentIn - $currentOutAll

    $balance = $balance + $currentBalance

    $expectedIncome = $expectedIncomeRate * $currentOutAll

    $expectedPriceForIncome = [System.Math]::Round(($expectedIncome + $currentOut) / (1 - $commissionRate) / $itemsCount, 2)

    $toDisplay = "$itemName `t $itemsCount `t $currentOut `t $currentIn `t $currentBalance `t $balance `t $expectedIncome `t $expectedPriceForIncome `t $currentPrice"

    $selectedColor = $null

    if ($currentBalance -lt 0)
    {
        $selectedColor = "Red"
    }
    elseif ($currentBalance -gt 0) {
        if ($currentBalance -gt $expectedIncome) {
            $selectedColor = "Green"
        }
        else {
            $selectedColor = "DarkGreen"
        }
    }
    else {
        $selectedColor = $null
    }

    if ($selectedColor -ne $null) {
        Write-Host $toDisplay -ForegroundColor $selectedColor
    }
    else {
        Write-Host $toDisplay
    }
}                      