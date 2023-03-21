# Microsoft Exchange Version Table
![Really not Microsoft](https://img.shields.io/badge/Not%20Official-Not%20Microsoft-red)
![Really not Microsoft](https://img.shields.io/badge/Scraped-From%20Microsoft-yellow)
![Data Counter Badge](https://img.shields.io/github/search/xasz/exchange-version-table/data?label=Data%20Counter%20%28Test%29)
![Last Commit](https://img.shields.io/github/last-commit/xasz/exchange-version-table)

**This is not an official Microsoft site. This is to help people simplify there scripts.**

All infomration is scraped from the official Micorosft documentation and processed into different formats to handle the information via simple calls.

I came across this issue several times and there was no satisfiyying good solution. I want to provide a simple solution to check if the current version of an Microsoft Exchange Server installation is Up2Date or does have a current CU Version installed.

For now i am only scraped the docs once and did an upload. My plan is to automate this process to provide steady and always up2date solution.

> I am not sure if this is allowed by microsoft. I will ask for permission if i find a channel to actually do that. I have not found any suitable page or documentation on this.


## Example Usage: Do i have the current Exchange Version

```powershell
# Run in Exchange Management Shell
$exchangeVersion = (Get-Command Exsetup.exe | ForEach {$_.FileVersionInfo}).ProductVersion
$data = Invoke-RestMethod "https://xasz.github.io/exchange-version-table/data/data.json"
if( ($data.Releases  | Where-Object {$_.BuildNumberLong -like $exchangeVersion}).IsCurrent){
    Write-Host "Your Exchange is Up2Date" -ForegroundColor Green
}else{
    Write-Host "Shame on you - Please install the latest Exchange Update" -ForegroundColor Red
}
```

## Example Usage: What CU do i have

```powershell
# Run in Exchange Management Shell
$exchangeVersion = (Get-Command Exsetup.exe | ForEach {$_.FileVersionInfo}).ProductVersion
$data = Invoke-RestMethod "https://xasz.github.io/exchange-version-table/data/data.json"
($data.Releases  | Where-Object {$_.BuildNumberLong -like $exchangeVersion}).CU
```

## Data Resources
| Format | Resource |
| - | - |
| Json | [https://xasz.github.io/exchange-version-table/data/data.json](data/data.json) | 
| Markdown | [https://xasz.github.io/exchange-version-table/data/data.md](data/data.md) |

