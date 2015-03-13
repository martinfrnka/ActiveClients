﻿#Globalni konfiguracni parametry

#pocet paralelne probihajicich jobu pro otestovani pocitacu pingem a zjisteni uzivatele pres WMI
#20 jobu pro cca 150 PC, probehne za cca 50s
$batchCnt = 20

#nazev domeny, ktery se ma odstranit z prihlaseneho uzivatele
#napr. z NSZBRN\mfrnka se odstrani NSZBRN\
$domain = "NSZBRN\"

#pomocna funkce pro barevny vypis na konzoli
function Write-Host-Color([String]$Text, [ConsoleColor]$Color) {
    Write-Host $Text -Foreground $Color -NoNewLine
}

#alias pro vyse uvedenou funkci
Set-Alias whc Write-Host-Color

#funkce pro generovani html kodu podle zjistenych dat
# vstupni parametry"
# filename - cesta a jmeno vystupniho html souboru
# columns  - pozadovany pocet sloupcu ve vysledne tabulce
function Export-Html-Data($filename, $columns) {
    #otevreni vystupniho souboru
    $fstream = New-Object System.IO.StreamWriter($filename, $false,  [System.Text.Encoding]::UTF8);
    
    #zapis HTML hlavicky a zacatku tabulky
    $fstream.Write("
<html>
    <head>
     <meta http-equiv=`"refresh`" content=`"45`">
        <link rel= `"stylesheet`" type=`"text/css`" href=`"style.css`">        
    </head>
    <body>
        <table>
            <tr>");

    #zapis nazvu sloupcu podle zvoleneho poctu
    for ($i = 0; $i -lt $columns; $i++){
        $fstream.Write("<th>Počítač</th><th>Uživatel</th>");
    }
    $fstream.WriteLine("</tr>");

    #kolik bude mit tabulka radku?
    $tbl_rows = [Math]::Ceiling($computer.Count / $columns);

    #cyklus pro zapis vsech radku tabulky
    for ($i = 0; $i -lt $tbl_rows; $i++) {

        #rozliseni lichych a sudych radku
        if (($i % 2) -eq 0) {
            $rozliseni="even-row";
        }
        else {
            $rozliseni="odd-row";
        }

        #hlavicka radku
        $fstream.WriteLine("<tr id=$rozliseni>");

        #generovani sloupcu aktualniho radku
        for ($c = 0; $c -lt $columns; $c++) {

            #kde v seznamu vysledku jsou data pro aktualni bunku?
            $idx = $i + $tbl_rows*$c;
            
            #pokud jsme se octli mimo seznam, bunka bude prazdna        
            if ($idx -ge $computer.Count) {
                $fstream.WriteLine("<td></td><td></td>");
                continue;
            }

            #je zapisovane PC spusteno nebo ne?
            if ($computerState[$idx][1] -eq $true) {
                $isActive=" id=active";
            } else {
                $isActive=" id=passive";
            }
            
            #vygenerovani html dat pro aktualni bunku
            $fstream.WriteLine("<td$isActive>"+$computer[$idx].ToString()+"</td><td$isActive>"+$computerState[$idx][2].ToString()+"</td>");

        }
        #ukonceni radku
        $fstream.WriteLine("</tr>");
    }
    #ukonceni tabulky a hlavicka tabulky se statistikami
    $fstream.Writeline("</table>
<table>
<tr>
<td><table>");
    #radky tabulky se statistikami a konec tabulky
    $date = Get-date;
    $fstream.Writeline("<tr id=`"even-row`"><td id=`"pc-total`">PC Celkem:</td><td id=`"pc-total`">$PC_Total</td></tr>");
    $fstream.Writeline("<tr id=`"odd-row`"> <td id=`"pc-on`">PC zapnuto:</td><td id=`"pc-on`">$PC_ON</td></tr>");
    $fstream.Writeline("<tr id=`"even-row`"><td id=`"pc-off`">PC vypnuto:</td><td id=`"pc-off`">$PC_OFF</td></tr>");
    $fstream.Writeline("<tr id=`"odd-row`"> <td id=`"users`">Přihlášených uživatelů:</td><td id=`"users`">$User_logged</td></tr>");
    $fstream.Writeline("<tr id=`"even-row`"><td id=`"time`">Čas zpracování:</td><td id=`"time`">$stop</td></tr>");
    $fstream.Writeline("<tr id=`"odd-row`"><td id=`"time`">Vygenerováno:</td><td id=`"timeGenerated`">$date</td></tr>");
    $fstream.Writeline("</table></td>
<td><table>
<tr id=`"even-row`"><td id=`"pc-off`">Stáří vygenerované sestavy:</td><td id=`"writeHere`"></td></tr>
</table></td>
</tr>
</table>");

    #javascript pro zjisteni rozdilu mezi aktualnim casem a casem vytvoreni
    #skript dopisuje rozdil v sekundach do tabulky statistik
    #podle toho kolik sekund od vygenerovani probehlo pouzije ruzne barvy
    #zelena do 120s, zluta do 600s, cervena nad 600s
    $fstream.Writeline("<script>
  var myVar=setInterval(function(){myTimer()},10);

  function myTimer() {
    var t_generated = new Date(document.getElementById(`"timeGenerated`").innerHTML);
   
    var t_now = new Date();
    var t_diff = Math.round((t_now.getTime() - t_generated.getTime())/1000);
    if (t_diff > 600) {
    document.getElementById(`"writeHere`").innerHTML = `"<span id=timeElapsedProblem>`" + t_diff.toString() + `" sekund.</span>`";
    } else if (t_diff > 120){
    document.getElementById(`"writeHere`").innerHTML = `"<span id=timeElapsedWarning>`" + t_diff.toString() + `" sekund.</span>`";
    } else {
    document.getElementById(`"writeHere`").innerHTML = `"<span id=timeElapsedOK>`" + t_diff.toString() + `" sekund.</span>`";
    }
  }
</script>

");

    #konec html dokumentu a uzavreni souboru
    $fstream.Writeline("</body></html>");
    $fstream.Close();

}

#scriptblock obsahuje skript, ktery je spousten v ramci jednotlivych jobu
#Vstupni parametr $computerNames obsahuje seznam pocitacu k otestovani
#Provadi nasledujici:
#    ping na pocitac $name vybrany z $computernames a zaznamenani vysledku
#    pokud pocitac odpovida, provede WMI dotaz na uzivatele a vysledek se zaznamena do pole $results
#Vraci pole $results, ve kterem je pro kazdy testovany pocitac zaznam o trech polozkach - nazev, stav, uzivatel
#    stav je true nebo false, (true = dostupny, false = nedostupny)
#    uzivatel obsahuje
#        - username
#        - '-----' v pripade ze je PC zapnuty a neni na nem nikdo prihlasen
#        - '' v pripade vypnuteho pocitace
$scriptblock = { 
    param (
        [string[]] $computerNames
    )
    #Write-Host "scriptblock pro: $computerNames"
    #Write-Host $computerNames.Count

    #priprava pole pro vysledky
    $results = @("")*$computerNames.Count

    #cyklus pres vsechny jmena pocitacu ze vstupu
    for ($i = 0; $i -lt $computerNames.Count; $i++)
    { 
        #vezmeme jmeno pocitace
        $name = $computerNames[$i]

        #testujeme ping
        if ((Test-Connection -ComputerName "$name" -Count 1 -Size 1 -Quiet))
        {
            #pokud je ping OK, pak se pres WMI pokusime ziskat prihlaseneho uzivatele 
            [string]$user = [string](Get-WMIObject -class Win32_ComputerSystem -ComputerName "$($name)").UserName
            if (($user -ne $null) ) 
            {
                #poznamename si informaci o uzivateli
                $user = $user.ToString().Replace($domain,"")
                if ($user.ToString().Length -eq 0) { $user = "-----" }
            }

            #do vysledku zapiseme zjistene veci o spustenem PC
            $results[$i] = @($name,$true,$user)
        }
        else
        { 
            #pocitac neodpovida, zapiseme do vysledku jako neaktivni
            $results[$i] = @($name,$false,"")
        }
    }

    #vratime vysledky
    return $results 
}


### Hlavni telo scriptu

#nekonecny cyklus
while ($true)
{
    $start = Get-Date

    $computer = Get-ADComputer -Filter * | ?{-not $_.Enabled -like 'f*'} | select name | %{$_.name} | sort
    #Write-Host "Pocetr PC: $($computer.count)"
    $computersInBatch = [Math]::Ceiling($computer.Count / $batchCnt)

    #priprava pole pro vysledky
    $computerState = @($false) * $computer.Count
    $userState = @("") * $computer.Count
    $jobs = @("") * $batchCnt


    #invokace jobu
    for ($i = 0; $i -lt $batchCnt;$i++) 
    {
   
        #Write-Host $name ": " -NoNewline
        #Write-Host (Invoke-Command -ScriptBlock $scriptblock -ArgumentList $name)
        #Set-Variable "ping$i" -Value {Invoke-Command -ScriptBlock $scriptblock -ArgumentList $computer[$i] -AsJob}
        $names = $computer[($i*$computersInBatch) .. (($i+1)*$computersInBatch-1)]
        #$names
        $jobs[$i] = Start-Job -name "ping$i" -ScriptBlock $scriptblock -ArgumentList (,$names) 
        #Write-Host "Job ping$i pro pocitace $names spusten"   
    

    }

    #Write-Host "cekam na provedeni procesu"
    Wait-Job -Name "ping*"
    #Write-Host "Procesy dokonceny"

    #stahni vysledky a odstran joby
    $computerState = $null
    for ($i = 0; $i -lt $batchCnt; $i++)
    { 
        $result = Receive-Job -Name "ping$i"
        $computerState += $result
        Remove-Job -Name "ping$i"    
    }

    Clear-Host
    Write-Host ("-"*119) 

    [int32]$colCnt = 4;
    [int32]$rowCnt = [Math]::Ceiling($computerState.Count/$colCnt) 

    $PC_ON = 0
    $PC_OFF = 0
    $User_logged = 0

    for ($row = 0; $row -lt $rowCnt; $row++)
    { 
        for ($col = 0; $col -lt $colCnt; $col++)
        { 
            $idx = ($col * $rowCnt) + $row
            if ($idx -ge $computerState.Count) 
            { 
                Write-Host (" "*28) -NoNewline
                whc "| " blue
            if ($col -eq $colCnt-1) 
            {
                Write-Host ""
            }

                continue 
            }

            if ($computerState[$idx][2] -eq $null) {
                $computerState[$idx][2] = "null";
            }

            if ($computerState[$idx][1] -eq $true)
            {
                $PC_ON++
                [string]$resultString = [string]::Format("{0,-14}{1,-14}",$computerState[$idx][0], " "+$computerState[$idx][2])
                whc $resultString green

                if (-not [string]::Equals($computerState[$idx][2].ToString(), "-----"))
                {
                    $User_logged++
                }

            }
            else
            {
                $PC_OFF++
                [string]$resultString = [string]::Format("{0,-28}",$computerState[$idx][0])
                whc $resultString gray
                #if (-not ($i%4 -eq 3)) {whc "| " blue}
            }

        
    
            if ($col -ne $colCnt-1) 
            {
                whc "| " blue
            }
            else
            {
                whc "|" blue
                Write-Host ""
            }


    
        }
    }
    $PC_Total = $PC_ON + $PC_OFF
 
    Write-Host ("`n" + "-"*119) 

    Write-Host ""
    Write-Host ""
    whc ("="*30) white
    Write-Host ""
    whc "PC Celkem:              $PC_Total`n" white
    whc "PC zapnuto:             $PC_ON`n" green
    whc "PC vypnuto:             $PC_OFF`n" gray
    whc "Prihlasenych uzivatelu: $User_logged`n" yellow
    whc ("="*30) white
    Write-Host ""


    $stop = Get-Date
    $stop -= $start
    $time = $stop.Second
    whc "Cas zpracovani: $stop " gray
    

    Export-Html-Data "E:\Temp\ActiveUsers4.html" 4;
    Export-Html-Data "E:\Temp\ActiveUsers5.html" 5;
    Export-Html-Data "E:\Temp\ActiveUsers.html" 6;
    Export-Html-Data "E:\Temp\ActiveUsers7.html" 7;
    Export-Html-Data "E:\Temp\ActiveUsers8.html" 8;

    $command = "C:\pracovni\winscp\WinSCP.com /script=C:\pracovni\winscp\sendfile.txt"
    iex $command

    Write-Host "`n"
    Write-Host "                <                    >`r" -NoNewline
    Write-Host "pauza 2 minuty  <" -NoNewline
    for ($i = 0; $i -lt 20; $i++)
    { 
        Start-Sleep -s 6
        Write-Host "*" -NoNewline
        
    }
    
    Write-Host "`rObnova dat, cca 50sekund...             ";


}
