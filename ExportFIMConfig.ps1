#========================================================================                 
# Created by:                   Henry Schleichardt / Thomas Herbst
# Organization:                 infoWAN Datenkommunikation GmbH
# Filename:                     ExportFIMConfig.ps1
# Script Version:               1.0
#========================================================================

#----------------------------------------------------------------------------------------------------------
 Set-Variable -Name URI   -value "http://localhost:5725/resourcemanagementservice" 
 Set-Variable -Name objectPrefix -Value "_BHOLD_"
 Set-Variable -Name xmlFile -Value "C:\Users\infowan\Desktop\FimConfig.xml"

 $global:allowModifyListOfGUIDS = $true
 $global:listOfGUIDS = @{}
 $date = Get-Date –f "yyyy-MM-dd HH:mm:ss"
 $xmlHeader = "<?xml version='1.0' encoding='utf-8'?>"
 $xmlHeader += "<!-- <content stamp='$date' file='$xmlFile' host='$env:computername' username='$env:username' /> -->"
 $xmlText = "$xmlHeader"


#----------------------------------------------------------------------------------------------------------


#--- Start Function Definition --------------------------------------------------------------------------------------------

Function addToListOfGUIDs
{
	Param(
        [parameter(Mandatory=$true, ValueFromPipeline = $true)]
        [String]
        $listEntry
    )
    Process
    {
		#  <ResourceCurrentSet>urn:uuid:8887df8e-6e84-49f2-a794-f9e9802077e0</ResourceCurrentSet>
		if ($global:allowModifyListOfGUIDS){
			if(-not $global:listOfGUIDS.ContainsKey($listEntry)){
				$global:listOfGUIDS.Add($listEntry,$null)
			}
		}
	
	}
}

Function Get-AttributeType
{
	Param
	(
	 	[parameter(Mandatory=$true, ValueFromPipeline = $true)]
		[Microsoft.ResourceManagement.Automation.ObjectModel.ResourceManagementAttribute] $obj
	)
	Process
	{			
		$psObject = New-Object PSObject
		
		if($obj.IsMultiValue.Equals($true))
		{
			$psObject | Add-Member -MemberType NoteProperty -Name "IsMultiValue" -Value $true
		} 
		else 
		{
			$psObject | Add-Member -MemberType NoteProperty -Name "IsMultiValue" -Value $false
		}		
		
		$value = $null
		if ($_.Value -ne $null)
        {
            $value = $_.Value
        } 
		else {
			if ($_.Values -ne $null)
			{
				$value = $_.values[0]
			}
		}
		
		if ($value -ne $null)
		{
			$psObject | Add-Member -MemberType NoteProperty -Name "IsBooleanValue" -Value  (($value -eq $true) -or ($value -eq $false))
			$psObject | Add-Member -MemberType NoteProperty -Name "IsObjectID" -Value $value.contains("urn:uuid")
			
			[datetime]$a = New-Object DateTime
			$result = [datetime]::tryparse($value, [ref]$a)
			$psObject | Add-Member -MemberType NoteProperty -Name "IsDateTime" -Value $result
			
			[int]$i = New-Object Int
			$result = [int]::TryParse($value, [ref]$i)
			$psObject | Add-Member -MemberType NoteProperty -Name "IsNumber" -Value $result
			
			if ( $psObject.IsBooleanValue -eq $true) 
			{
				$psObject | Add-Member -MemberType NoteProperty -Name $_.AttributeName -Value $_.Value
			} else 
			{
				$psObject | Add-Member -MemberType NoteProperty -Name $_.AttributeName -Value $_.Value				
			}
			
		} else {
			$psObject | Add-Member -MemberType NoteProperty -Name $_.AttributeName -Value $null
		}		
		
		Write-Output $psObject		
	}
}

Function Convert-FimExportToXMLObject
{
    Param
    (
        [parameter(Mandatory=$true, ValueFromPipeline = $true)]
        [Microsoft.ResourceManagement.Automation.ObjectModel.ExportObject]
        $ExportObject
    )
    Process
    {  
		$xmlString = ""
        $psObject = New-Object PSObject
        $ExportObject.ResourceManagementObject.ResourceManagementAttributes | ForEach-Object{
		
			$attributeType = Get-AttributeType($_)
			if ($attributeType.IsMultiValue -eq $true) {
				$values = $_.values
				$append = "<" + $_.AttributeName
				if ($_.AttributeName.Contains("Workflow")){ $append += " Type =" + [char]34 + "WorkflowDefinition"  + [char]34}
				$append += " >`n"
				foreach($value in $values){ 
					$temp = $value
					if ($attributeType.IsObjectID -eq $true)
					{
						$temp = $value.split(':')[2]
						addToListOfGUIDs($temp)
					}
					$append += "<Value>" + $temp + "</Value>`n"
				}
				$append += "</" + $_.AttributeName + ">`n"
				$xmlString += $append
			} 
			else 
			{
				if ($attributeType.IsObjectID -eq $true){
					$temp = $_.value.split(':')[2]
					addToListOfGUIDs($temp)
					$append = "<" + $_.AttributeName + ">" + $temp +  "</" + $_.AttributeName + ">`n"
				} 
				else
				{
					$append = "<" + $_.AttributeName  
					if (-not ($attributeType.IsBooleanValue -eq $null)) { $append += " Boolean = " + [char]34 + $attributeType.IsBooleanValue.ToString() + [char]34 }
					if (-not ($attributeType.IsNumber -eq $null)) { $append += " Number = " + [char]34 + $attributeType.IsNumber.ToString() + [char]34 }
					if (-not ($attributeType.IsDateTime -eq $null)) { $append += " DateTime = " + [char]34 + $attributeType.IsDateTime.ToString() + [char]34 }
					$append += " >" + $_.value
					$append +=  "</" + $_.AttributeName + ">`n"
				}
				$xmlString += $append
			}
			
			if ($_.AttributeName -eq "DisplayName"){
				$psObject | Add-Member -MemberType NoteProperty -Name "xmlNodeName" -Value $_.Value
			}
			if ($_.AttributeName -eq "ObjectID"){
				$temp = $_.value.split(':')[2]
				$psObject | Add-Member -MemberType NoteProperty -Name "xmlNodeID" -Value $temp
			}
			
			
            $psObject | Add-Member -MemberType NoteProperty -Name $_.AttributeName -Value $append
			
        }
		$psObject | Add-Member -MemberType NoteProperty -Name "XMLString" -Value $xmlString
        Write-Output $psObject
    }
} 


#--- End Function Definition --------------------------------------------------------------------------------------------


 if(@(get-pssnapin | where-object {$_.Name -eq "FIMAutomation"} ).count -eq 0) {add-pssnapin FIMAutomation}
 


#----------------------------------------------------------------------------------------------------------
# Get all MPRs  
#----------------------------------------------------------------------------------------------------------
#
#  $allMpr = export-fimconfig -uri "http://localhost:5725/resourcemanagementservice" `
#                             -customconfig ("/ManagementPolicyRule") `
#							  –onlyBaseResources `
#							  -ErrorVariable Err `
#							  -ErrorAction SilentlyContinue

#  $allMpr = export-fimconfig -uri "http://localhost:5725/resourcemanagementservice" `
#                             -customconfig ("/ManagementPolicyRule[Disabled = 'False']") `
#							  –onlyBaseResources `
#							  -ErrorVariable Err `
#							  -ErrorAction SilentlyContinue

  $allMpr = export-fimconfig -uri "http://localhost:5725/resourcemanagementservice" `
                              -customconfig ("/ManagementPolicyRule[(starts-with(DisplayName,'$($objectPrefix)')) and (Disabled = 'False')]") `
							  –onlyBaseResources `
							  -ErrorVariable Err `
							  -ErrorAction SilentlyContinue
  
  $xmlText += "<ManagementPolicyRules>"
  foreach ($mpr in $allMpr){
    #$temp = "<MPR Name = " + [char]34 + $myMPR.xmlNodeName  + [char]34 + ">`n"	
	$myMPR =  $mpr |  Convert-FimExportToXMLObject
	$xmlText += "<MPR Name = " + [char]34 + $myMPR.xmlNodeName  + [char]34 + " Id =  " + [char]34 + $myMPR.xmlNodeID  + [char]34 + ">`n" 
	$xmlText += $myMPR.XMLString
	$xmlText += "</MPR>"
  }
  
  $global:allowModifyListOfGUIDS = $false
 
  $xmlText += "<SETs>"
 	foreach($GUID in $global:listOfGUIDS.Keys){
   
 	$theSet = export-fimconfig -uri "http://localhost:5725/resourcemanagementservice" `
                              -customconfig ("/Set[ObjectID='$($GUID)']") `
							  –onlyBaseResources `
							  -ErrorVariable Err `
							  -ErrorAction SilentlyContinue
							  
	if(-not($theSet -eq $null)){
		$mySet = $theSet | Convert-FimExportToXMLObject
		$xmlText += "<SET Name = " + [char]34 + $mySET.xmlNodeName  + [char]34 + " Id =  " + [char]34 + $mySET.xmlNodeID  + [char]34 + ">`n" 
		$xmlText += $mySET.XMLString
		$xmlText += "</SET>"
	}
	
  }
  $xmlText += "</SETs>"
  
  $xmlText += "<Workflows>"
  foreach($GUID in $global:listOfGUIDS.Keys){
   
 	$theWorkflow = export-fimconfig -uri "http://localhost:5725/resourcemanagementservice" `
                              -customconfig ("/WorkflowDefinition[ObjectID='$($GUID)']") `
							  –onlyBaseResources `
							  -ErrorVariable Err `
							  -ErrorAction SilentlyContinue
							  
	if(-not($theWorkflow -eq $null)){
		$myWorkflow = $theWorkflow | Convert-FimExportToXMLObject
		$xmlText += "<Workflow Name = " + [char]34 + $myWorkflow.xmlNodeName  + [char]34 + " Id =  " + [char]34 + $myWorkflow.xmlNodeID  + [char]34 + ">`n" 
		$xmlText += $myWorkflow.XMLString
		$xmlText += "</Workflow>"
	}
	
  }
  
  $xmlText += "</Workflows>"
  
 $xmlText += "</ManagementPolicyRules>"
 [xml]$xmlDoc = $xmlText
 $xmlDoc.Save($xmlFile)


							  