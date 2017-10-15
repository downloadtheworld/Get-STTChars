$rarities = "Common","Uncommon","Rare","Super_Rare","Legendary"
#$rarities = "Uncommon"
$file = '.\STTChars.xlsx'
$newchars = @()
foreach ($rarity in $rarities){
    $characters = Invoke-WebRequest https://stt.wiki/wiki/Category:$rarity -UseBasicParsing
    #$characters.links.href

    $begin = ""
    $end = ""
    if($rarity -eq "Common"){
        $begin = "/wiki/Cadets"
        $end = "https://stt.wiki/w/index.php?title=Category:Common&amp;oldid=104908"
    }
    if($rarity -eq "Uncommon"){
        $begin = "/wiki/Cadets"
        $end = "https://stt.wiki/w/index.php?title=Category:Uncommon&amp;oldid=105597"
    }
    if($rarity -eq "Rare"){
        $begin = "#p-search"
        $end = "https://stt.wiki/w/index.php?title=Category:Rare&amp;oldid=105227"
    }
    if($rarity -eq "Super_Rare"){
        $begin = "#p-search"
        $end = "https://stt.wiki/w/index.php?title=Category:Super_Rare&amp;oldid=105562"
    }
    if($rarity -eq "Legendary"){
        $begin = "#p-search"
        $end = "https://stt.wiki/w/index.php?title=Category:Legendary&amp;oldid=105122"
    }

    $begun = 0
    $chars = @()
    foreach ($link in $characters.Links){
        if ($link.href -eq $end) {
            break
        }
        if ($begun -eq "1"){
            $char = New-Object –TypeName PSObject
            $char | Add-Member –MemberType NoteProperty -Name "Name" -Value $link.title
            $char | Add-Member –MemberType NoteProperty -Name "Link" -Value $link.href
            $char | Add-Member –MemberType NoteProperty -Name "Rarity" -Value $rarity
            $chars += $char
        
        }
        if ($link.href -eq $begin) {
            $begun = 1
        }
    }
    $chars.Count
    $newchars += $chars
}
$olddata = Import-Excel $file

foreach ($oldchar in $olddata){
    foreach ($newchar in $newchars){
        if ($oldchar.name -eq $newchar.name){
            $newchar | Add-Member –MemberType NoteProperty -Name "Active" -Value $oldchar.Active -Force
            $newchar | Add-Member –MemberType NoteProperty -Name "Frozen" -Value $oldchar.Frozen -Force
            $newchar | Add-Member –MemberType NoteProperty -Name "Immortalised" -Value $oldchar.Immortalised -Force
        }
    }
    
}

$newchars | Export-Excel $file
$newchars.Count
