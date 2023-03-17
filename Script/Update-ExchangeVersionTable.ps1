Import-Module PowerHTML


$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$MS_EXCHANGE_VERSION_PAGE = "https://docs.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates"
$Response = Invoke-WebRequest -Uri $MS_EXCHANGE_VERSION_PAGE
$html = ConvertFrom-Html -Content $Response.Content

$filter = @('Product name', 'Release date', 'Build number(short format)', 'Build number(long format)')


$VersionTables = @()
$evHtmlTables = $html.SelectNodes('//table') | Where-Object { -not(Compare-Object $_.ChildNodes['thead'].ChildNodes['tr'].Elements('th').InnerText $filter) }
foreach($evHtmlTable in $evHtmlTables){
    $isFirst = $true
    $versions = @()
    $MajorReleaseName = $null

    foreach($tr in $evHtmlTable.Element('tbody').SelectNodes('tr')){
        $data = $tr.Elements('td').InnerText.Trim()
        if($data.Count -ne 4 -or $data[0].Length -eq 0) { continue }
        if(!$MajorReleaseName) { $MajorReleaseName = ($data[0] | Select-String -Pattern "Exchange Server \d*").Matches[0].Value }
        $realDate = $null
        try {
            $realdate = Get-Date $data[1]
            Write-Verbose "Could convert $($data[1]) into a DateTime on $($data[0])"
        }catch{
            Write-Verbose "Could not convert $($data[1]) into a DateTime$($data[0])" -ForegroundColor Yellow
        }
        $versions += [PSCustomObject]@{
            ProductName = $data[0]
            Date = $realDate
            BuildNumberShort = $data[2]
            BuildNumberLong = $data[2]
            Newest = $isFirst
        }
        $isFirst = $false
    }
    if($MajorReleaseName){
        $VersionTables += [PSCustomObject]@{
            MajorReleaseName = $MajorReleaseName
            MicrosoftVersionListing = $versions
            NewestRelease = $versions | Where-Object {$_.Newest -eq $true}
        }
    }else{
        Write-Verbose "Could not find MajorRelease - Ignoring Content"
    }
}

