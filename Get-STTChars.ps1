#For when there is more than 500 in a category:

#Raw information about the members of a category, their sortkeys and timestamps (time when last added to the category) can be obtained from the API, using a query of the form:

#http://en.wikipedia.org/w/api.php?cmtitle=Category:Category_name&action=query&list=categorymembers&cmlimit=500&cmprop=title|sortkey|timestamp
#Listings of up to 500 members are possible. If there are more members then the results will include text near the end like this: <categorymembers cmcontinue="page|NNNN|TITLE" />.

#This can be added to the previous one, without quotation marks, for the next page of members: ...&cmcontinue=page|NNNN|TITLE








$rarities = "Common","Uncommon","Rare","Super_Rare","Legendary"
#$rarities = "Uncommon"
$file = '.\STTChars.xlsx'
$newchars = @()
foreach ($rarity in $rarities){
    $request = 'https://stt.wiki/w/api.php?cmtitle=Category:' + $rarity + '&action=query&list=categorymembers&cmlimit=500&cmprop=title|sortkey|timestamp&format=json'
$request
    $json = Invoke-WebRequest $request
    $json
    $converted = ConvertFrom-Json -InputObject $json
    $converted.query.categorymembers
    Write-Host "converted"

    $chars = @()
    foreach ($link in $converted.query.categorymembers){

        $char = New-Object –TypeName PSObject
        $char | Add-Member –MemberType NoteProperty -Name "Name" -Value $link.title
        $char | Add-Member –MemberType NoteProperty -Name "Rarity" -Value $rarity
        $chars += $char
    }
    #$chars.Count
    $newchars += $chars
}
if (Test-Path $file){
    $olddata = Import-Excel $file

    foreach ($oldchar in $olddata){
        foreach ($newchar in $newchars){
            if ($oldchar.name -eq $newchar.name){
                $newchar | Add-Member –MemberType NoteProperty -Name "Active" -Value $oldchar.Active -Force
                $newchar | Add-Member –MemberType NoteProperty -Name "Immortalised" -Value $oldchar.Immortalised -Force
                $newchar | Add-Member –MemberType NoteProperty -Name "Frozen" -Value $oldchar.Frozen -Force
            }
        }
    
    }
}
$newchars | Export-Excel $file -BoldTopRow -AutoFilter -FreezeTopRow
$newchars.Count
