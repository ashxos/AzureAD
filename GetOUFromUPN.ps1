$Report = import-csv "UPNUserList.csv"

$GroupedReport = $Report | group -Property UPN

$items =0

foreach ( $Row in $GroupedReport )
{
    $items++
    #$UPN = (Get-EXOMailbox -Identity  $Row.Name).UserPrincipalName
	#$PrimarySMTP = (Get-EXOMailbox -Identity  $Row.Name).PrimarySmtpAddress
	#$PrimarySMTP = (Get-EXOMailbox -UserPrincipalName $Row.Name).PrimarySmtpAddress
	#$UserId = (Get-AzureADUser -Filter "userPrincipalName eq '\"$Row.Name\"'").ObjectId
	$UserId = (Get-AzureADUser -Searchstring $Row.Name).ObjectId 

	Write-host "$UserId"
	$OU = (Get-AzureADUserExtension -ObjectId $UserId).onPremisesDistinguishedName
	Write-host "$OU"
    Write-Progress -Activity "Getting UPN of user" -Status "$items of $($GroupedReport.Count)" -PercentComplete ( ($items/$GroupedReport.Count)*100)
    Write-host "$items of $($GroupedReport.Count)"
    foreach ( $Entry in $Row.group)
    {  
        $Entry | Add-Member -NotePropertyName 'OU' -NotePropertyValue $OU
		#$Entry | Add-Member -NotePropertyName 'PrimarySmtpAddress' -NotePropertyValue $PrimarySMTP
		#$Entry | Add-Member -NotePropertyName 'UserPrincipalName' -NotePropertyValue $UPN
    }
}

$GroupedReport.group | export-csv 'UPN_OU.csv' -NoTypeInformation -Encoding UTF8
