Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

#filters for all the enabled AAFC users
$TestGroup = Get-ADUser -Filter "Company -eq 'some_company'" | Where-Object { $_.enabled -eq $True }

$regAboutMeEmpty = 0
$regAboutMeNonEmpty = 0
$abtMeNonEmpty = 0
#the province empty about me hashtable
$provEmptyAbt = @{ }

#the province non empty about me hashtable    
$provNonEmptyAbt = @{ }

#adds the provinces to hashtable
function province_counter {
    param([hashtable] $provinces, [string] $prov)
    if (!$provinces.ContainsKey($prov)) {
        $provinces.Add($prov, 1)
    }
    else {
        $provinces[$prov]++
    }
}

function Remove-StringLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}

#$serviceContext = Get-SPServiceContext -Site "http://blahblahblah.com"
#$profileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($serviceContext)
$servercontext = [Microsoft.Office.Server.ServerContext]::Default
$profilemanager = new-object Microsoft.Office.Server.UserProfiles.UserProfileManager($servercontext)
#$profiles = $profileManager.GetEnumerator()

#loops through all the enabled profiles
foreach ($user in $TestGroup.SamAccountName) {
    
    try {
        #retrieves the profile from profile manager
        $full_user = 'some_company\' + $user
        $prof = $profilemanager.GetUserProfile($full_user)


        #sets the province string and the about me string
        $str_abt = $prof["AboutMe"].value
        $str_prov = $prof["Province"].value
    
        #checks number of about me filled out
        if ($str_abt) {
            $abtMeNonEmpty++      
        } 

        #checks if the province field has been filled out
        if ($str_prov) {
            
            $str_prov = Remove-StringLatinCharacters -String $str_prov
            $str_prov = $str_prov.ToUpper().Replace(".", "").Replace("-", "").Replace(" ", "").Replace("|", "")
            
        
            #checks which province 
            switch ($str_prov) {
                { $_.IndexOf("ONT") -ne -1 } { $result = "ON"; break }
                { $_.IndexOf("ALBERTA") -ne -1 } { $result = "AB"; break }
                { $_.IndexOf("BRITISHCOLUMBIA") -ne -1 } { $result = "BC"; break }
                { $_.IndexOf("MANITOBA") -ne -1 } { $result = "MB"; break }
                { $_.IndexOf("NEWBRUNSWICK") -ne -1 } { $result = "NB"; break }
                { $_.IndexOf("NEWFOUNDLANDANDLABRADOR") -ne -1 } { $result = "NL"; break }                 
                { $_.IndexOf("NOVASCOTIA") -ne -1 } { $result = "NS"; break }
                { $_.IndexOf("NORTHWEST") -ne -1 } { $result = "NT"; break }
                { $_.IndexOf("NUNAVUT") -ne -1 } { $result = "NU"; break }
                { $_.IndexOf("PEI") -ne -1 -Or $_.IndexOf("PRINCEEDWARDISLAND") -ne -1 } { $result = "PE"; break }
                { $_.IndexOf("QUEBEC") -ne -1 } { $result = "QC"; break }
                { $_.IndexOf("SASKATCHEWAN") -ne -1 } { $result = "SK"; break }
                { $_.IndexOf("YUKON") -ne -1 } { $result = "YT"; break }
                default { $result = $_ }
            }
    
            #checks how many provinces filled out also have about me               
            if ($str_abt) {
                $regAboutMeNonEmpty++
                province_counter -provinces $provNonEmptyAbt -prov $result      
            }
             
            #checks how many provinces filled out have no about me 
            else {
                $regAboutMeEmpty++
                province_counter -provinces $provEmptyAbt -prov $result
            }
        }     

    }
    catch {
        Write-Warning "User '$user' does not exist. Skipping..."
    }

}
      
$abtMeEmpty = $TestGroup.SamAccountName.count - $abtMeNonEmpty
$provEmptyAbt
write-host "---------------------"
$provNonEmptyAbt
    
<#-------------------------------Output Section------------------------------#>
    
[PSCustomObject]@{
    'Name'           = "Total Active users AboutMe not empty"
    'Count'          = $abtMeNonEmpty
    'Region'         = "Canada"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Total Active users AboutMe empty"
    'Count'          = $abtMeEmpty
    'Region'         = "Canada"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $regAboutMeNonEmpty
    'Region'         = "Unkown"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $regAboutMeEmpty
    'Region'         = "Unkown"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["ON"] 
    'Region'         = "ON"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["ON"] 
    'Region'         = "ON"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["AB"]
    'Region'         = "AB"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["AB"]
    'Region'         = "AB"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["BC"]
    'Region'         = "BC"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["BC"]
    'Region'         = "BC"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["MB"]
    'Region'         = "MB"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["MB"]
    'Region'         = "MB"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["NB"] 
    'Region'         = "NB"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["NB"]
    'Region'         = "NB"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["NL"]
    'Region'         = "NL"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["NL"]
    'Region'         = "NL"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["NS"]
    'Region'         = "NS"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["NS"]
    'Region'         = "NS"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["NT"]
    'Region'         = "NT"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["NT"]
    'Region'         = "NT"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["NU"]
    'Region'         = "NU"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["NU"]
    'Region'         = "NU"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["PE"]
    'Region'         = "PE"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["PE"]
    'Region'         = "PE"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["QC"]
    'Region'         = "QC"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["QC"]
    'Region'         = "QC"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["SK"]
    'Region'         = "SK"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["SK"]
    'Region'         = "SK"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe not empty"
    'Count'          = $provNonEmptyAbt["YT"] 
    'Region'         = "YT"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
}, [PSCustomObject]@{
    'Name'           = "Regional Active users AboutMe empty"
    'Count'          = $provEmptyAbt["YT"] 
    'Region'         = "YT"
    'Processed Time' = Get-Date -Format "MM/dd/yyyy HH:mm"
} | Export-Csv AbtMePrototype.csv -NoTypeInformation 







         
