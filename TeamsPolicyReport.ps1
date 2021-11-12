
#Functions
Function PivotPolicyData 
{
    Param ($Data, $PolicyURI)
    

    $Members = $Data | get-member -MemberType NoteProperty | Where {$_.Name -ne "Identity" -and $_.Name -ne "Description" -and $_.Name -ne "Key"}  
  

    $Result = New-Object System.Data.DataTable
    $Result.Columns.Add("Policy Property")
    $Result.Columns.Add("Property Description")
    $Data | % {$Result.Columns.Add($_.Identity)} | Out-Null
    $descriptions = GetPolicyPropertyDesc -PolicyURI $PolicyURI
    $Members |  % {
    $row = $Result.NewRow(); 
    $row.'Policy Property' = $_.Name; 
    $PolicyProp = $_.Name      
    $description = $descriptions | Where {$_.Property -like "-$PolicyProp"} | Select-Object -ExpandProperty Description

    $row.'Property Description' = $description

    foreach($Policy in $Data)
        {     
           $PolicyID = $Policy.Identity
           $Value = $Policy | Select-Object -ExpandProperty $PolicyProp
           $row[$PolicyID] = $Value 
        } 

    $Result.Rows.Add($row)
    } | Out-Null
    return $Result | Format-Table   
    
}


Function GetPolicyPropertyDesc 
{
Param ($PolicyURI)

$URI = $PolicyURI
$HTML = Invoke-WebRequest -Uri $URI
$Elements = $HTML.AllElements | where-object {$_.Class -like "parameter*"} | Select-Object innerText
$PropDesc = @()
foreach($index in (1..($Elements.Count -1)))
{
$PropDesc += [pscustomobject]@{Property = $Elements[$index - 1].innerText; Description = $Elements[$index].innerText}
}
return $PropDesc
}





Connect-MicrosoftTeams

$MeetingPolicy = Get-CsTeamsMeetingPolicy | Select-Object -Property * -ExcludeProperty XsAnyElements,XsAnyAttributes,Element,CommaSeparator,pscomputerName,RunspaceId,PSShowComputerName 
$MeetingPolicyPiv = PivotPolicyData -Data $MeetingPolicy -PolicyURI "https://docs.microsoft.com/en-us/powershell/module/skype/set-csteamsmeetingpolicy?view=skype-ps"
$MeetingPolicyPiv[0].Table | Export-csv -Path "C:\Users\Documents\GIT Repos\MS Teams\Meeting Policy Evaluation\TeamsMeetingPolicies.csv" -NoTypeInformation


$CallPolicy = Get-CSTeamsCallingPolicy | Select-Object -Property * -ExcludeProperty XsAnyElements,XsAnyAttributes,Element,CommaSeparator,pscomputerName,RunspaceId,PSShowComputerName 
$CallPolicyPiv = PivotPolicyData -Data $CallPolicy -PolicyURI "https://docs.microsoft.com/en-us/powershell/module/skype/set-csteamscallingpolicy?view=skype-ps"
$CallPolicyPiv[0].Table | Export-csv -Path "C:\Users\Documents\GIT Repos\MS Teams\Meeting Policy Evaluation\CallPolicies.csv" 

