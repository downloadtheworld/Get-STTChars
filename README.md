# Get-STTChars
Powershell script that grabs Star Trek Timelines characters names and rarities from the wiki and writes them to a spreadsheet. Mostly written because I like to play with Powershell sometimes.

Requires the ImportExcel Powershell module from https://github.com/dfinke/ImportExcel. Only tested on Windows 10 with Powershell 5.

This script will write out to a spreadsheet file called STTChars.xlsx in the same folder as the script. This includes columns for Active, Immortalised and Frozen. This will stay in the sheet even when new characters are added. It also includes rarity and link, which will revert if changed.

# Known issues
No error handling. You might want to backup your STTChars file before re-running the script.

The Link field isn't an actual link but you can work out how to use it.
