Function download($url, $name) {
    Write-Host -NoNewline "Downloading $url ...... "
    wget $url -Method GET -OutFile $name -ErrorAction SilentlyContinue
    Write-Host -ForegroundColor Green "Succeed." 
}

cd (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
cd html
mkdir GuildLine,Summary,Balance,Profit,Dupont -ErrorAction Ignore
cd ..
$sinaRoot = "http://vip.stock.finance.sina.com.cn/corp/go.php"
$guildHead = $sinaRoot +"/vFD_FinancialGuideLine/stockid/"
$summaryHead = $sinaRoot + "/vFD_FinanceSummary/stockid/"
$balanceHead = $sinaRoot + "/vFD_BalanceSheet/stockid/"
$profitHead = $sinaRoot + "/vFD_ProfitStatement/stockid/"
$dupontHead = $sinaRoot + "/vFD_DupontAnalysis/stockid/"

$guildTail = "/displaytype/4.phtml"
$summaryTail = "/displaytype/4.phtml"
$balanceTail = "/ctrl/part/displaytype/4.phtml"
$profitTail = "/ctrl/part/displaytype/4.phtml"
$dupontTail = "/displaytype/10.phtml"

$idList = gc "conf/idList.txt"

# Download
foreach($id in $idList) {

    $url = $guildHead + $id + $guildTail
    download $url "html/GuildLine/$id.html"

    $url = $summaryHead + $id + $summaryTail
    download $url "html/Summary/$id.html"

    $url = $balanceHead + $id + $balanceTail
    download $url "html/Balance/$id.html"

    $url = $profitHead + $id + $profitTail
    download $url "html/Profit/$id.html"

    $url = $dupontHead + $id + $dupontTail
    download $url "html/Dupont/$id.html" 

}

# Analyze
Add-Type -Path "lib\HtmlAgilityPack.dll"
Add-Type -Path "lib\ClosedXML.dll"
copy "data/template.xlsx" “temp.xlsx” -Force
$workbook = New-Object ClosedXML.Excel.XLWorkbook(“temp.xlsx”)

$guildLineHeaderList = (gc "conf/guildLineHeader.txt" | Select-String -NotMatch "#").Line
$headerCount = $guildLineHeaderList.Length
$worksheet = $workbook.Worksheet(1)
$row = 2
$col = 3
foreach($header in $guildLineHeaderList) {
    $worksheet.Cell($row, $col++).Value = $header
    $worksheet.Cell($row, $col + $headerCount).Value = $header
}

$doc = New-Object HtmlAgilityPack.HtmlDocument
foreach($id in $idList) {
    # Id
    $worksheet.Cell(++$row,1).Value = $id

    # GuildLine
    Write-Host -NoNewline "Processing $id ...... "
    $doc.Load("html/GuildLine/$id.html")
    $table = $doc.DocumentNode.SelectNodes("//table[@id='BalanceSheetNewTable0']")
    if($table -eq $null) {
        Write-Host "Failed" -ForegroundColor Red
        # Retry Download
        $url = $guildHead + $id + $guildTail
        download $url "html/GuildLine/$id.html"
        $doc.Load("html/GuildLine/$id.html")
        $table = $doc.DocumentNode.SelectNodes("//table[@id='BalanceSheetNewTable0']")
        if($table -eq $null) {
            Write-Host "Page loading failure, please check." -ForegroundColor Red
            continue;
        }
    }
    # Name
    $worksheet.Cell($row, 2).Value = $table.InnerText.Substring(0,$table.InnerText.IndexOf('(')).Trim()

    $col = 3;
    foreach($header in $guildLineHeaderList) {
        $tr = $table.SelectNodes("//tr[contains(td, '$header')]")
        $worksheet.Cell($row, $col++).Value = $tr.ChildNodes[1].InnerText
        $worksheet.Cell($row, $col + $headerCount).Value = $tr.ChildNodes[2].InnerText
    }
    Write-Host -ForegroundColor Green "Succeed." 
}

$workbook.SaveAs("财务报表.xlsx")
del "temp.xlsx" -Force -ErrorAction Ignore
