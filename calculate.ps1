

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

    $data = Invoke-WebRequest $url

    $receivedData = ConvertFrom-Json -InputObject $data

    $bidValue = $receivedData.BidQuotes[$receivedData.BidQuotes.Length - 1][1]

    $askValue = $receivedData.AskQuotes[$receivedData.AskQuotes.Length - 1][1]

    $refValue = $receivedData.ReferenceQuotes[$receivedData.ReferenceQuotes.Length - 1][1]

    Write-Host "[Get-IngItemData] $itemName $bidValue $askValue $refValue"

    $result = ($bidValue, $askValue, $refValue)

    $result
}

$data = Import-csv -Path ".\..\export.csv" -Delimiter ';'

$cashOut = 0.0
$cashIn = 0.0
$commission = 0.0
$commissionRate = 0.0039
$balance = 0.0
$expectedIncomeRate = 0.05

foreach ($item in $data)
{
    $itemName = $item."Papier"

    Write-Host "Item $itemName"

    if ($itemName.StartsWith("INTL"))
    {
        $currentOut = ([double]$item."Kurs K") * ([int]$item."Pozycja")

        $resultValues = Get-IngItemData $itemName

        $currentCommissionOut = $currentOut * $commissionRate

        $currentIn = $resultValues[0] * ([int]$item."Pozycja")

        $currentCommissionIn = $currentIn * $commissionRate

        $currentOutAll = $currentOut + $currentCommissionOut + $currentCommissionIn

        $currentBalance = $currentIn - $currentOutAll

        $balance = $balance + $currentBalance

        $expectedIncome = $expectedIncomeRate * $currentOutAll

        $toDisplay = "$itemName `t $currentOut `t $currentIn `t $currentBalance `t $balance `t $expectedIncome"

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