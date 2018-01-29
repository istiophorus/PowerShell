
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

$data = Import-csv -Path "c:\My\Scripts\export.csv" -Delimiter ';'

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

    if ($itemName.StartsWith("INTL"))
    {
        $itemsCount = [int]$item."Pozycja"

        $currentOut = ([double]$item."Kurs K") * ($itemsCount)

        $resultValues = Get-IngItemData $itemName

        $currentCommissionOut = $currentOut * $commissionRate

        $currentPrice = $resultValues[0]

        $currentIn = [System.Math]::Round($currentPrice * $itemsCount, 2)

        $currentCommissionIn = [System.Math]::Round($currentIn * $commissionRate, 2)

        $currentOutAll = [System.Math]::Round($currentOut + $currentCommissionOut + $currentCommissionIn, 2)

        $currentBalance = $currentIn - $currentOutAll

        $balance = $balance + $currentBalance

        $expectedIncome = $expectedIncomeRate * $currentOutAll

        $expectedPriceForIncome = [System.Math]::Round(($expectedIncome + $currentOut) / (1 - $commissionRate) / $itemsCount, 2)

        $incomePercent = [System.Math]::Round($currentBalance * 100.0 / ($currentOut + $currentCommissionOut), 2)

        $toDisplay = "$itemName `t $itemsCount `t $currentOut `t $currentIn `t $currentBalance `t $balance `t $expectedIncome `t $expectedPriceForIncome `t $currentPrice `t $incomePercent %" 

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
    elseif ($itemName.StartsWith("RC"))
    {
        # not implemented yet
    }
}                      