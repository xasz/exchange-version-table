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
    
    $latestCU = $null
    
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
        $cuMatches = ($data[0] | Select-String -Pattern "Cu(\d+)").Matches
        
        $cuVersion = ""
        if($cuMatches.Groups.Count -eq 2){
            $cuVersion = $cuMatches.Groups[1].Value
            if($null -eq $latestCU){
                $latestCU = $cuVersion
            }
        }

        $versions += [PSCustomObject]@{
            ProductName = $data[0]
            Date = $realDate
            BuildNumberShort = $data[2]
            BuildNumberLong = $data[3]
            CU = $cuVersion
            CUOffset = $latestCU - $cuVersion
            IsCurrent = $isFirst
        }
        $isFirst = $false
    }
    if($MajorReleaseName){
        $VersionTables += [PSCustomObject]@{
            MajorReleaseName = $MajorReleaseName
            MicrosoftVersionListing = $versions
            CurrentRelease = $versions | Where-Object {$_.IsCurrent -eq $true}
        }
    }else{
        Write-Verbose "Could not find MajorRelease - Ignoring Content"
    }
}




# Json Export
$objectData = [PSCustomObject]@{
    GenerationDate = Get-Date
    Source = $MS_EXCHANGE_VERSION_PAGE
    Releases = $VersionTables.MicrosoftVersionListing
}
$objectData | ConvertTo-Json | Out-File -FilePath "data/data.json" 

# Markdown Export
$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine("# Exchange Version Table");
[void]$sb.AppendLine("Generation: $((Get-Date).ToString())");
[void]$sb.AppendLine("Source: $MS_EXCHANGE_VERSION_PAGE");
[void]$sb.AppendLine();

foreach($VT in $VersionTables){
    [void]$sb.AppendLine("# {0}" -f $VT.MajorReleaseName);
    [void]$sb.AppendLine();
    [void]$sb.AppendLine("| Product Name | Date | Build Number (Short) | Build Number (Long) |");
    [void]$sb.AppendLine("|:- | - | - | - |");
    foreach($V iN $VT.MicrosoftVersionListing){
        if($V.IsCurrent -eq $true){
            [void]$sb.AppendLine(("| **{0}** | **{1}** | **{2}** | **{3}** |" -f $V.ProductName, $V.Date.ToString("yyyy-MM-dd"), $V.BuildNumberShort, $V.BuildNumberLong));
        }else{
            [void]$sb.AppendLine(("| {0} | {1} | {2} | {3} |" -f $V.ProductName, $V.Date.ToString("yyyy-MM-dd"), $V.BuildNumberShort, $V.BuildNumberLong));
        }
    }            
    [void]$sb.AppendLine();
}

$readmeTeamplate += $sb.ToString() | Out-File -FilePath "data/data.md"


# Append to Readme
$readmeTeamplate = Get-Content ".\templates\Readme.md"
$readmeTeamplate += $sb.ToString()
$readmeTeamplate | Out-File -FilePath ".\Readme.md"
