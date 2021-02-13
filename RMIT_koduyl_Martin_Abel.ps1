#Skript on valminud RMIT-i kodutöö raames kandideerimaks MS Süsteemiadministraatatori kohale.
#Skript on loodud Martin Abel-i poolt 13.02.2021

#Impordime AD ja Exchange Powershelli moodulid.
Import-Module activedirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

#Skriptis kasutame veahaldust - Vigu kuvatakse try - catch meetodil"
$ErrorActionPreference = 'silentlycontinue'

#Logifail luuakse kui script jõuab esimese funktsiooni kus seda kasutatakse, siin määrame ära, et logifail oleks unikaalne"
$logiFail = ".\log_"+$(Get-Date -Format "dd-MM-yyyy_hh-mm-ss")+".txt"
$loodudKasutajadLog = ".\loodud_kasutajad_"+$(Get-Date -Format "dd-MM-yyyy_hh-mm-ss")+".txt"

#Globlaased muutujad, mida scriptis kasutatakse
$domeen = "rmit.ee"
$ADServer = "ad.domeen.intra"
$EXCServer = "exch.domeen.intra"
$OU="OU=Users,OU=RMIT,DC=DOMEEN,DC=INTRA"
$EmailDataBase = "RMIT"
$ABP = "RMIT"
$ASMP = "RMIT"

#Testime kas AD serveriga on ühendus. Testitakse 3 korda, kui ei õnnestu väljastatakse veateade ning skript lõpetab töö
try {
    Test-NetConnection -ComputerName $ADServer -Hops 3 -WarningAction Stop > $null
}
catch {
    Write-Output "$ADServer ei saadud ühendust"
    Exit
    }
#Testime kas Exchange serveriga on ühendus. Testitakse 3 korda, kui ei õnnestu väljastatakse veateade ning skript lõpetab töö
try {
    Test-NetConnection -ComputerName $EXCServer -Hops 3 -WarningAction Stop > $null
    }
catch {
    Write-Output "$EXCServer ei saadud ühendust"
    Exit
    }

#Funktsioonis muudab ettantud tähed, funktsioonis kontrollitakse, et skripti on etteantud  sisendfail
function Get-CSVFix($sourceCSV, $fixedCSV){
    try {
        ((Get-Content -Path $sourceCSV -ErrorAction Stop).Replace("ğ","s").Replace("ş","z").Replace("õ","o").Replace("ö","o").Replace("ä","a").Replace("ü","y").Replace("Ğ","S").Replace("Ş","Z").Replace("Õ","O").Replace("Ä","A").Replace("Ö","O").Replace("Ü","Y")) | Set-Content -Path $fixedCSV

    }
    catch{
        Write-Output "Faili $sourceCSV ei leitud"    
    }
}

#Funktsioon otsib muutuja sisule sünonüümi hashtabelist. Kui sünonüüm ekisteerib, asendab ta muutuja sisu.
function Convert-RMITOsakond($osaKond){
  $okond = New-Object -TypeName HashTable

  $okond.'Personaliosakond' = 'PO'
  $okond.'Haldusosakond' = 'HO'

    
  Foreach ($key in $okond.Keys)
  {
    $osaKond = $osaKond.Replace($key, $okond.$key)
  }
  $osaKond
}

#Funktsioon loob kasutajanime. Kasutajanime puhul kasutatakse RMIT-i standardit ees.perenimi. Kasutajanimi genereeritakse eesnime ning perenimi muutujast.
#Kasutaja parool genereertiakse isikukoodist, kus lisatakse sõnale "Parool" loodava kasutaja isikukoodist 4 viimast numbirt.
#Kasutaja parool muudetakse Secure-Stringiks ja kasutatakse konto loomisel.
function Loo-Kasutaja($eesnimi, $perenimi, $ik, $osakond){
	$kasutajaNimi = [string]::Format("{0}.{1}", $eesnimi, $perenimi)
	$neliViimast = "$ik".Substring("$ik".Length -4)
	$PlainPw = [string]::Format("{0}{1}", "Parool", $neliViimast)
	$SecurePw = ConvertTo-SecureString $PlainPw -AsPlainText -Force

#Kasutaja luuakse kui seda ei eksisteeri. Kasutaja olemasolukorral kuvatakse info logifailis.
#Kasutaja loomisel kontrollitakse, et ei esineks vigu. Kui viga esineb salvestatakse see logifaili ning jätkatakse skriptiga.
	
    if (Get-ADuser -F {SamAccountName -eq $kasutajaNimi}){
		Write-Output "Kasutaja $kasutajaNimi eksisteerib" >> $logiFail
	}

	else{

#Kasutaja loomisel kontrollitakse, et ei esineks vigu. Kui viga esineb lisatakse see logifaili. Skript jätkab tööd.

        try{
		    New-ADUser -SamAccountName $kasutajaNimi -UserPrincipalName "$kasutajaNimi@$domeen" -Enabled $False -ChangePasswordAtLogon $True -Name "$eesnimi $perenimi" -GivenName "$eesnimi" -Surname "$perenimi" -DisplayName "$eesnimi $perenimi" -AccountPassword $SecurePw -Path $OU -ErrorAction Stop
            Write-OutPut "$kasutajaNimi loodi parooliga $PlainPw" >> $loodudKasutajadLog
            }
        catch{
            Write-Output "$kasutajaNimi ei eksisteeri kuid mingil põhjusel ei õnnestunud seda luua" >> $logiFail
            {Continue}>$null
            }

#Kasutajale isikukoodi lisamisel kontrollitakse, et ei esineks vigu. Kui viga esineb lisatakse see logifaili. Skript jätkab tööd.

		try{
            Set-ADUser -Identity $kasutajaNimi -Add @{personalcode="$ik"} -ErrorAction Stop
            }
        catch{
            Write-Output "$kasutajaNimi ei saanud IK-d lisada" >> $logiFail
            {Continue}>$null
            }
            
#Kasutaja gruppi lisamisel kontrollitakse, et ei esineks vigu. Kui viga esineb lisatakse see logifaili. Skript jätkab tööd.
#Grupi nimi valitakse sisendfailist. Esmalt muudetaks see ümber nagu hastabels.	
		try{
            Add-ADGroupMember -Identity (Convert-RMITOsakond -osaKond $osakond) -Members $kasutajaNimi -ErrorAction Stop
            }
        catch{
            Write-Output "$kasutajaNimi ei õnnestunud osakonna gruppi lisada" >> $logiFail
            {Continue}>$null
            }
		
	}
}

#Funktsioon loob emaili kui seda ei eksisteeri. Kontrollitakse SMTP olemasolu nii mail useritel kui ka kontaktidel.

function Tee-Email($eesnimi, $perenimi){
	$kasutajaNimi = [string]::Format("{0}.{1}", $eesnimi, $perenimi)
	if (Get-Recipient $kasutajaNimi@$domeen ){
		Write-Output "$kasutajaNimi@$domeen SMTP eksisteerib ning seda postkasti ei loodud" >> $logiFail
	}

	else{

#Kui SMTP puudub, luuakse kasutajale mailbox. Kontrollitakse kas mailboxi loomine õnnestub. Vea korral kuvatakse viga logifailis.
        try{
		    Enable-Mailbox -Identity $kasutajaNimi -Database "$EmailDataBase" -DisplayName "$eesnimi $perenimi" -PrimarySmtpAddress "$kasutajaNimi@$domeen" -Alias $kasutajaNimi -AddressBookPolicy $ABP -ActiveSyncMailboxPolicy $ASMP -ErrorAction Stop > $null
	    }
        catch{
            Write-Output "$kasutajaNimi@$domeen ei eksisteeri, aga mingil põhujsel ei suudetud seda luua" >> $logiFail
            {Continue}>$null
        }
    }
}

#Kasutatakse eelpool kirjeldatud funktsiooni Get-CSV, et muuta CSV sisu RMIT-i kasutajate loomise standardile vastavaks.
Get-CSVFix -sourceCSV .\uuedkasutajad.csv -fixedCSV .\standard_uuedkasutajad.csv

#Loetakse parandatud CSV sisse.
$ADUsers = Import-Csv .\standard_uuedkasutajad.csv

#CSV-st käiakse kõik kasutajad läbi ning proovitakse luua kasutaja ning email kasutades eelkirjeldatud funktsioone Loo-Kasutaja ning Tee-Email.
foreach ($User in $ADUsers){
	Loo-Kasutaja -eesnimi $User.Eesnimi -perenimi $User.Perenimi -ik $User.Isikukood -osakond $User.Osakond
	Tee-Email -eesnimi $User.Eesnimi -perenimi $User.Perenimi
}
#Scripti töö lõppedes kuvatakse logifailid.

#Logifail kui on kirjeldatud vead mis skripti tööd ei peatanud + (kasutajad ning emailid mida ei loodud).
Get-Content $logiFail

#Logifail kus on välja toodud kasutajad mis loodi ning loodud kasutajate paroolid.
Get-Content $loodudKasutajadLog