 <#
.SYNOPSIS 
    API CONNEXION SCRIPT 
.DESCRIPTION
	https://iadlife.peoplespheres.com/
.NOTES
	Authors : 
		Jeremie Poujol (jeremie.poujol@iadinternational.com)
	Version : 
		V1.0 January 2022
.LINK
        https://fdocuments.net/document/api-documentation-guide.html?page=1
        Old documentation: https://rest.monportailrh.com/swagger/ 
        New documentation: https://rest.monportailrh.com/docs/ 

.EXAMPLE
    [ps] E:\Powershell\Scheduled_task\API_PeopleSphere.ps1
#>

Clear-Host

$logo = @"
.__            .___ ___.                .__                                _____               __                       
|__|____     __| _/ \_ |__  __ __  _____|__| ____   ____   ______ ______ _/ ____\____    _____/  |_  ___________ ___.__.
|  \__  \   / __ |   | __ \|  |  \/  ___/  |/    \_/ __ \ /  ___//  ___/ \   __\\__  \ _/ ___\   __\/  _ \_  __ <   |  |
|  |/ __ \_/ /_/ |   | \_\ \  |  /\___ \|  |   |  \  ___/ \___ \ \___ \   |  |   / __ \\  \___|  | (  <_> )  | \/\___  |
|__(____  /\____ |   |___  /____//____  >__|___|  /\___  >____  >____  >  |__|  (____  /\___  >__|  \____/|__|   / ____|
        \/      \/       \/           \/        \/     \/     \/     \/              \/     \/                   \/     

"@
Write-Host $logo -ForegroundColor "green"
Write-Host "Bienvenue dans iadlife.peoplespheres.com" -ForegroundColor "white"
Start-Sleep -Seconds 1

# Check password support-itps@iadinternational.com exist 
While (!(Test-Path C:\Scripts\SecureString\support-itps@iadinternational.com.$env:username.securestring))
{
    Write-Host "No SecureString has been found for user" $env:username "in C:\Scripts\SecureString folder" -ForegroundColor Yellow
    read-host "Enter support-itps@iadinternational.com's Password" -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath C:\Scripts\SecureString\support-itps@iadinternational.com.$env:username.securestring
}

# Authentication - To get a Bearer token, you should use the API request below:
$url = "https://sso.monportailrh.com/auth/realms/Internal-idp/protocol/openid-connect/token"

# Token expires within 5 minutes, during this time you can: 
#  1. Login again 
#  2. Or, Refresh the token by sending following request:
<#
• POST https://sso.monportailrh.com/auth/realms/Internal-idp/protocol/openid-connect/token 
• Header: "Content-Type" = application/x-www-form-urlencoded 
• Body: 
o grant_type: refresh_token 
o client_id: realm-management 
o refresh_token: {the expired token} 
#>

$startTime = Get-Date

<#  In a PowerShell script that is currently running, you can use the StartTime property of the Get-Process object to get the time when the script was started, and then use the Elapsed property of the Stopwatch object to get the elapsed time since the script began running. You can then use an if statement to check if this time is greater than 4 minutes and, if so, execute a specific action. Here is an example code that can help you:
$stopwatch = [diagnostics.stopwatch]::StartNew()
# your code here
# check if script has been running for more than 4 minutes
if ($stopwatch.Elapsed -gt [timespan]::FromMinutes(4)) {
    # execute your specific action here
}
#>

$secureString = Get-Content "C:\Scripts\SecureString\support-itps@iadinternational.com.$env:username.securestring" | ConvertTo-SecureString
$clearText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString))

<#
$password = Read-Host "Please enter password" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
$SecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
#> 
  
$body = [ordered]@{
    "grant_type" = "password";
    "client_id" = "monportailrh-web-app";
    "username" = "support-itps@iadinternational.com";
    "password" = $clearText
}  
$header = [ordered]@{
    "authority" = "rest.monportailrh.com";
    "sec-ch-ua" = "Not A;Brand`";v=`"99`", `"Chromium`";v=`"96`", `"Google Chrome`";v=`"96`"";
    "accept" = "application/json, text/plain, */*";
    "accept-language" = "fr";
    "sec-ch-ua-mobile" = "?0";
    "user-agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36";
    "sec-ch-ua-platform" = "Windows";
    "origin" = "https://app.monportailrh.com";
    "sec-fetch-site" = "same-site";
    "sec-fetch-mode" = "cors";
    "sec-fetch-dest" = "empty";
    "referer" = "https://app.monportailrh.com/"
}

$response = Invoke-RestMethod -Uri $url -Method 'Post' -Body $body -Headers $header
$bearer_access_token = $response.access_token
$token=1
<#
Get Users
Get users per pagination, the request below is to get the first 25 active users
GET     https://rest.monportailrh.com/search?name=&pso-type=usr&page=1&active=1&per-page=25 

• pso-type=usr 
• page=1 
• active=1 
• per-page=25
#>

$url2 = "https://rest.monportailrh.com/search?include=data.quick_actions&name=&page=1&per-page=25&pso-type=usr"
$header2 = [ordered]@{
    "accept" = "application/json, text/plain, */*";
    "authorization" = "Bearer $bearer_access_token"
}
$result2 = Invoke-RestMethod -Uri $url2 -Headers $header2
# sample : $result2.data.data[0].fields
$totalusers = $result2.data.meta.pagination.total

write-host $totalusers "Utilisateurs a traiter dans PeopleSphere" -ForegroundColor Red -BackgroundColor Yellow

$url3 = "https://rest.monportailrh.com/search?name=&pso-type=usr&page=1&active=1&per-page=$totalusers"
$result3 = Invoke-RestMethod -Uri $url3 -Headers $header2

##############################################################################
#$result3.data.data | Where-Object {$_.id -eq 10}
#$result3.data.data | Where-Object {$_.fields.value -eq "poujol"}
#$result3.data.data | Where-Object { $_.fields | Where-Object { $_.id -eq 125 -and $_.value -eq "POUJOL" } }
##############################################################################

<#
$id = 484
$url4 = "https://rest.monportailrh.com/psos/$id/fields?active=true&include=type,items,options,settings,assignment_settings"
$result4 = Invoke-RestMethod -Uri $url4 -Headers $header2
foreach ($item in $result4.data) {
    Write-Output "ID: $($item.id) - Name: $($item.name) - Alias: $($item.alias)"
}
Write-Output $result4.data[0]
#>

$result=@()
foreach ($id in $result3.data.data.id) {
    $url4 = "https://rest.monportailrh.com/psos/$id/fields?active=true&include=type,items,options,settings,assignment_settings"
    $result4 = Invoke-RestMethod -Uri $url4 -Headers $header2
    Write-host -ForegroundColor yellow "--"
    Write-Host -ForegroundColor yellow "Traitement user id #$id - " $result4.data[9].value_details #mail pro si la position dans le tableau de bouge pas /!\
    Write-host -ForegroundColor yellow "--"
    $ids = 711, 2637, 683, 2638, 1145, 125, 123, 423, 550, 28, 27, 182, 344, 185, 186, 348, 859
    <# $result4.data | Select-Object id, name
        id name
        -- ----
        711 Date de début dans le poste
        2637 [YOUZER] Date d'embauche
        683 Départ de la société
        2638 - [YOUZER] Date de sortie
        1145 Type collaboration
        125 Nom
        123 Prénom
        423 Image de profil
        550 Civilité
        28 Adresse e-mail professionnelle
        27 Téléphone portable professionnel
        182 Responsable
        344 Service
        185 Poste
        186 Entité Légale
        348 Site
        859 Matricule
    #>
    <#  support-itps@iadinternational.com ()
        New ID found not expected : 665 - Numéro de téléphone personnel
        New ID found not expected : 666 - Adresse e-mail personnelle
        New ID found not expected : 648 - Adresse personnelle
        New ID found not expected : 511 - Contact en cas d'urgence
        New ID found not expected : 525 - Personne à charge
        New ID found not expected : 516 - Situation familiale
        New ID found not expected : 667 - Expérience professionnelle
        New ID found not expected : 746 - Formation
        New ID found not expected : 735 - Certifications
        New ID found not expected : 1073 - LinkedIn
        New ID found not expected : 680 - Langue
        New ID found not expected : 1077 - Attestation employeur
        New ID found not expected : 1103 - Attestation Mutuelle Obligatoire
        New ID found not expected : 1108 - Attestation non perception supplément familial
        New ID found not expected : 1125 - Demande d'attestation
        New ID found not expected : 127 - Nom d'utilisateur
        New ID found not expected : 626 - Mutuelle
        New ID found not expected : 612 - Compte bancaire
    #>
    foreach ($item in $result4.data) {
        if ($ids -contains $item.id) {
            Write-Host -foregroundColor blue "$($item.id) - $($item.name) : " -NoNewline
            if ($item.id -eq 711) {
                $datededebutdansleposte = $item.value_details
                Write-Host $datededebutdansleposte
            }
            elseif ($item.id -eq 2637) {
                $youzerdatedembauche = $item.value_details
                Write-Host $youzerdatedembauche
            }
            elseif ($item.id -eq 683) {
                $departdelasociete = $item.value_details
                Write-Host $departdelasociete
            }
            elseif ($item.id -eq 2638) {
                $youzerdatedesortie = $item.value_details
                Write-Host $youzerdatedesortie
            }
            elseif ($item.id -eq 1145) {
                $typecollaboration = $item.value_details
                Write-Host $typecollaboration
            }
            elseif ($item.id -eq 125) {
                $nom = $item.value_details
                Write-Host $nom
            }
            elseif ($item.id -eq 123) {
                $prenom = $item.value_details
                Write-Host $prenom
            }
            elseif ($item.id -eq 423) {
                $imagedeprofil = $item.value_details
                Write-Host $imagedeprofil
            }
            elseif ($item.id -eq 550) {
                $civilite = $item.value_details
                Write-Host $civilite
            }
            elseif ($item.id -eq 28) {
                $adresseemailprofessionnelle = $item.value_details
                Write-Host $adresseemailprofessionnelle 
            }
            elseif ($item.id -eq 27) {
                $telephoneportableprofessionnel = $item.value_details
                Write-Host $telephoneportableprofessionnel
            }
            elseif ($item.id -eq 182) {
                $Responsable = $item.value_details.professional_email
                Write-Host $Responsable
            }
            elseif ($item.id -eq 344) {
                $Service = $item.value_details
                Write-Host $Service
            }
            elseif ($item.id -eq 185) {
                $Poste = $item.value_details
                Write-Host $Poste
            }
            elseif ($item.id -eq 186) {
                $entitelegale = $item.value_details
                Write-Host $entitelegale
            }
            elseif ($item.id -eq 348) {
                $site = $item.value_details
                Write-Host $site
            }
            elseif ($item.id -eq 859) {
                $matricule = $item.value_details
                Write-Host $matricule
            }
        } else {
            <# Action when all if and elseif conditions are false #>
            Write-Host -foregroundColor Red "New ID found not expected : $($item.id) - $($item.name)"
        }        
    }
    $hash=[ordered]@{
        "Date de début dans le poste"=$datededebutdansleposte;
        "[YOUZER] Date d'embauche"=$youzerdatedembauche;
        "Départ de la société"=$departdelasociete;    
        "[YOUZER] Date de sortie"=$youzerdatedesortie
        "Type collaboration"=$typecollaboration;
        "Nom"=$nom;
        "Prénom"=$prenom;
        "Image de profil"=$imagedeprofil;
        "Civilité"=$civilite;
        "Adresse e-mail professionnelle"=$adresseemailprofessionnelle;
        "Téléphone portable professionnel"=$telephoneportableprofessionnel;
        "Responsable"=$Responsable;
        "Service"=$Service;
        "Poste"=$Poste;
        "Entité Légale"=$entitelegale;
        "Site"=$site;
        "Matricule"=$matricule;
        "DateExeScript"=Get-Date;
    }
    $obj=New-Object psobject -property $hash
    $result+=$obj
    $elapsedTime = (Get-Date) - $startTime
    if ($elapsedTime.TotalMinutes -ge 4 -and $token -eq 1) {
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        write-host $totalusers "The token 1 is about to expire  - refresh_token for a second one - "$elapsedTime.TotalMinutes" minutes" -ForegroundColor Red -BackgroundColor Yellow
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        # Script has been running for 4 minutes or more, avoid expiration and refresh the token by sending following request
        $response = Invoke-RestMethod -Uri $url -Method 'Post' -Body $body -Headers $header
        $bearer_access_token = $response.access_token
        $header2 = [ordered]@{
            "accept" = "application/json, text/plain, */*";
            "authorization" = "Bearer $bearer_access_token"
        }
        $token=2           
    } elseif ($elapsedTime.TotalMinutes -ge 8 -and $token -eq 2) {
        # Code à exécuter si le temps écoulé est supérieur ou égal à 8 minutes et que $token est égal à 2
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        write-host $totalusers "The token 2 is about to expire  - refresh_token for a third one - "$elapsedTime.TotalMinutes" minutes" -ForegroundColor Red -BackgroundColor Yellow
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        # Script has been running for 8 minutes or more, avoid expiration and refresh the token by sending following request
        $response = Invoke-RestMethod -Uri $url -Method 'Post' -Body $body -Headers $header
        $bearer_access_token = $response.access_token
        $header2 = [ordered]@{
            "accept" = "application/json, text/plain, */*";
            "authorization" = "Bearer $bearer_access_token"
        }
        $token=3 # avoid to enter in the loop
    } elseif ($elapsedTime.TotalMinutes -ge 12 -and $token -eq 3) {
        # Code à exécuter si le temps écoulé est supérieur ou égal à 8 minutes et que $token est égal à 2
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        write-host $totalusers "The token 3 is about to expire  - refresh_token for a fourth one - "$elapsedTime.TotalMinutes" minutes" -ForegroundColor Red -BackgroundColor Yellow
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        # Script has been running for 12 minutes or more, avoid expiration and refresh the token by sending following request
        $response = Invoke-RestMethod -Uri $url -Method 'Post' -Body $body -Headers $header
        $bearer_access_token = $response.access_token
        $header2 = [ordered]@{
            "accept" = "application/json, text/plain, */*";
            "authorization" = "Bearer $bearer_access_token"
        }
        $token=4 # avoid to enter in the loop
    } elseif ($elapsedTime.TotalMinutes -ge 16 -and $token -eq 4) {
        # Code à exécuter si le temps écoulé est supérieur ou égal à 8 minutes et que $token est égal à 2
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        write-host $totalusers "The token 3 is about to expire  - refresh_token for a fourth one - "$elapsedTime.TotalMinutes" minutes" -ForegroundColor Red -BackgroundColor Yellow
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        # Script has been running for 12 minutes or more, avoid expiration and refresh the token by sending following request
        $response = Invoke-RestMethod -Uri $url -Method 'Post' -Body $body -Headers $header
        $bearer_access_token = $response.access_token
        $header2 = [ordered]@{
            "accept" = "application/json, text/plain, */*";
            "authorization" = "Bearer $bearer_access_token"
        }
        $token=5 # avoid to enter in the loop
    } elseif ($elapsedTime.TotalMinutes -ge 20 -and $token -eq 5) {
        # Code à exécuter si le temps écoulé est supérieur ou égal à 8 minutes et que $token est égal à 2
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        write-host $totalusers "The token 3 is about to expire  - refresh_token for a fourth one - "$elapsedTime.TotalMinutes" minutes" -ForegroundColor Red -BackgroundColor Yellow
        write-host -ForegroundColor Red -BackgroundColor Yellow "--"
        # Script has been running for 12 minutes or more, avoid expiration and refresh the token by sending following request
        $response = Invoke-RestMethod -Uri $url -Method 'Post' -Body $body -Headers $header
        $bearer_access_token = $response.access_token
        $header2 = [ordered]@{
            "accept" = "application/json, text/plain, */*";
            "authorization" = "Bearer $bearer_access_token"
        }
        $token=6 # avoid to enter in the loop
    } else {
        # normal operation, script continuity
    }
}

# Define the base path without extension (for the UTF-16 "production" export)
$basePath = "E:\Powershell\03-FlatFilesStorage\AzureSQLDatabase_csv\GenerateCSV-API-PeopleSpheres"

# Export to UTF-16 (Unicode) for Excel compatibility (production output)
$exportUnicode = "$basePath-excel.csv"
$result | Export-Csv -NoTypeInformation -Path $exportUnicode -Encoding Unicode

# Define target folder for the UTF-8 export
$utf8Folder = "E:\Powershell\03-FlatFilesStorage\GenerateCSV-API-PeopleSpheres\"
# Create the folder if it doesn't exist
if (-not (Test-Path $utf8Folder)) {
    New-Item -ItemType Directory -Path $utf8Folder | Out-Null
}

# Generate a timestamp for the UTF-8 filename
$timestamp = Get-Date -Format "yyyyMMdd"
$exportUtf8 = "$utf8Folder\GenerateCSV-API-PeopleSpheres-utf8-$timestamp.csv"

# Export to UTF-8 with timestamp (for automation/scripts)
$result | Export-Csv -NoTypeInformation -Path $exportUtf8 -Encoding UTF8

# Output confirmation
Write-Host "✅ CSV exports completed:"
Write-Host " - UTF-16 (Excel / production): $exportUnicode"
Write-Host " - UTF-8  (Script / timestamped): $exportUtf8"

########################
# UPLOAD FILE TO AZURE #
########################
$tenant = "e419a47d-b189-44f1-a28e-16be83c1f11e"
$subscription = "f9d75155-d6d0-4867-b0c0-cec83ecea40c"
$Usertenant = "AzSce_PSScript@iadgroup.onmicrosoft.com"

$ResourceGroupName = "rg-frc-coreservices"
$StorageAccountName = "iadsamgmtfrccore"  #iadstorageaccount122022

$ServerInstance = "sql-frc-coreservices-iad.database.windows.net"
$Database = "iaddb"
$Database_SA_LOGIN = "PSScript"

While (!(Test-Path C:\scripts\SecureString\$Usertenant.$env:username.securestring))
{
    read-host "Enter $Usertenant's Password" -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath C:\scripts\SecureString\$Usertenant.$env:username.securestring
}
$secure_passwd=Get-Content "C:\Scripts\SecureString\$Usertenant.$env:username.securestring" | ConvertTo-SecureString
$Credential = New-Object System.Management.Automation.PSCredential ($Usertenant , $secure_passwd)

#Connexion Azure dans le bon contexte
Connect-AzAccount -Credential $Credential -Tenant $tenant -Subscription $subscription

#contexte storage account
$storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -StorageAccountName $StorageAccountName
$context= $storageAccount.Context
$containerName = "csvcontainer"

#local file path et blob  
$localFilePath=$exportPath
$blobName="GenerateCSV-API-PeopleSpheres.csv" 

#upload blob
Set-AzStorageBlobContent -Container $containerName -Blob $blobName -File $localFilePath -Context $context -Force

While (!(Test-Path C:\Scripts\SecureString\$ServerInstance.$env:username.securestring))
{
    Write-Host "No SecureString has been found for user" $env:username "in C:\Scripts\SecureString folder" -ForegroundColor Yellow
    read-host "Enter SA Password for $ServerInstance" -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath C:\Scripts\SecureString\$ServerInstance.$env:username.securestring
}

$secString=Get-Content "C:\Scripts\SecureString\$ServerInstance.$env:username.securestring" | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PSCredential ($Database_SA_LOGIN,$secString)

$passwd=$credential.GetNetworkCredential().password

$query=@"
TRUNCATE TABLE dbo.PEOPLESPHERE_Iad
BULK INSERT dbo.PEOPLESPHERE_Iad
FROM 'csvcontainer/GenerateCSV-API-PeopleSpheres.csv'
WITH (DATA_SOURCE = 'blobcontainer', FIRSTROW = 2, FORMAT = 'CSV');
"@

$params = @{
   'Database' = $Database
   'ServerInstance' =  $ServerInstance
   'Username' = $Database_SA_LOGIN
   'Password' = $passwd
   'OutputSqlErrors' = $true
   'Query' = $query
   }

Invoke-Sqlcmd  @params -EncryptConnection

<#
# Enable  TLS 1.2 and Define password pour no-reply-ps
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$smtppw = Get-Content "C:\Scripts\SecureString\No-Reply-PS@iadinternational.com.$env:username.securestring" | ConvertTo-SecureString
$credentials = New-Object System.Management.Automation.PSCredential ("No-Reply-PS@iadinternational.com", $smtppw)
#$credentials = New-Object System.Management.Automation.PSCredential("no-reply-ps@iadinternational.com", (ConvertTo-SecureString "password" -AsPlainText -Force))
While (!(Test-Path C:\scripts\SecureString\No-Reply-PS@iadinternational.com.$env:username.securestring)){
    Write-Host "No SecureString has been found for user" $env:username "in C:\scripts\SecureString\" -ForegroundColor Yellow
    read-host "Enter No-Reply-PS@iadinternational.com's Password" -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath C:\scripts\SecureString\No-Reply-PS@iadinternational.com.$env:username.securestring
}

# Function signature mail
function Get-ScriptName
{ 
	return $MyInvocation.ScriptName; 
}


$body = "Nouveau csv generated" 
$ScriptName = $ScriptName | Get-ScriptName
$body+="<br><br>RunAs cov\$env:username On $env:computername <i>(local path : $ScriptName)</i><br>"
#$body+="Log file on bf-lie-mgt01 : $($LogFilePath)<br>"
$body+=Get-content -Path E:\Powershell\02_Common\email_signature\SignatureSI.txt


$emailParams = @{
    To = "jeremie.poujol@iadinternational.com"
    Cc = "jeremie.poujol@iadinternational.com"
    #No-Reply-PS - Network and System IT <986c2ea3.iadgroup.onmicrosoft.com@fr.teams.ms>
    Bcc = "jeremie.poujol@iadinternational.com"
    From = "no-reply-ps@iadinternational.com"
    Subject = "API CONNEXION SCRIPT  - iadlife - peoplesphere"
    Body = $body
    BodyAsHtml = $true
    Encoding = [System.Text.Encoding]::UTF8
    SmtpServer = "smtp.office365.com"
    Credential = $credentials
    Port = 587
    UseSsl = $true   
}
Send-MailMessage @emailParams
#>

<# SQL Server Management Studio (SSMS) Creation query

CREATE TABLE dbo.PEOPLESPHERE_Iad (
	Datededebutdansleposte VARCHAR(255),
    Youzerdatedembauche VARCHAR(255),
	Departdelasociete VARCHAR(255),
    Youzerdatedesortie VARCHAR(255),
    Typecollaboration VARCHAR(255),
	Nom VARCHAR(255),
	Prenom VARCHAR(255),
	Imagedeprofil VARCHAR(255),
	Civilite VARCHAR(255),
	Adresseemailprofessionnelle VARCHAR(255),
	Telephoneportableprofessionnel VARCHAR(255),
	Responsable VARCHAR(255),
    Service VARCHAR(255),
    Poste VARCHAR(255),
    Entitelegale VARCHAR(255),
    Site VARCHAR(255),
    Matricule VARCHAR(255),
	DateExeScript VARCHAR(255)
)

#>
