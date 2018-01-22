#========================================================================                 
# Created by:                   Thomas Herbst / Henry SchleichardtS
# Organization:                 infoWAN Datenkommunikation GmbH
# Filename:                     xmlToVisio.ps1
# Script Version:               1.0
#========================================================================


#region CustomVariables
	#Path to save the final PNG´s and the Visiodocument
	$savePath = "C:\Users\henry.schleichardt\Desktop\PNG"
	#Path for the FIMConfig.xml
	$filePath = "C:\Users\henry.schleichardt\Desktop\FimConfig.xml"
	#Path to the FIMVisio.vss custom stencils
	$customStencils = "C:\Users\henry.schleichardt\Documents\Meine Shapes\FIMVisio.vss"
#endregion

#region Constants

	#Set some constants for connection-direction between Visioshapes
	Set-Variable -Name Auto -Value "0" #-Option Constant
	Set-Variable -Name Down -Value "2" #-Option Constant
    Set-Variable -Name Left -Value "3" #-Option Constant
	Set-Variable -Name Right -Value "4" #-Option Constant	

#endregion 

#region variables and initilization

    #open Visio, create New Document and create a new Page
    $activeVisApp = Get-VisioApplication;
	
    $visApp = New-VisioApplication
	$visDoc = New-VisioDocument

    #Spacing between shapes
    $startRow = 0.0
    $startLine = 0.0
    
    $diffRow = 3.0
    $diffLine = 1.5
	
    #load the prefered stencil sets
	$basic_u = Open-VisioDocument basic_u.vss
    $connec_u = Open-VisioDocument connec_u.vss 
	$floshp_m = Open-VisioDocument floshp_m.vss
    
    #load our custom stencils
    $fimstencil = Open-VisioDocument $customStencils
    
    #rectangle behind requestor and targetbefore
    $rectShape1 = Get-VisioMaster "Master.2" $fimstencil
    #rectangle behind Workflows
    $rectShape2 = Get-VisioMaster "Master.3" $fimstencil
    #rectangle behind Workflow-Properties
    $rectShape3 = Get-VisioMaster "Master.4" $fimstencil
    #rectangle behind the whole thing
    $rectShape4 = Get-VisioMaster "Master.5" $fimstencil
    #Shape for legend on the right
    $legendShape = Get-VisioMaster "Master.11" $fimstencil
    #invisible Connector for the legend
    $legendConnector = Get-VisioMaster "Master.9" $fimstencil
    #ActionType and ActionParameter connector
    $actionConnector = Get-VisioMaster "Master.10" $fimstencil


    #load different shapes for mprs, sets, wfs...

	$mprShape = Get-VisioMaster "Framed Rectangle" $floshp_m
    $mprActionShape = Get-VisioMaster "Microform Processing" $floshp_m
    
   
    $setShape = Get-VisioMaster "Create Request" $floshp_m
    $setParameterShape = Get-VisioMaster "Microform Recording" $floshp_m
   
   
    $wfShape = Get-VisioMaster "Data Store 3" $floshp_m
    $wfActivityShape = Get-VisioMaster "File Of Cards" $floshp_m

    $connector = Get-VisioMaster "Dynamic Connector" $connec_u
#endregion

#region DrawFunctions

	function PinToRequestor
	{
		Param($source, $target, $shape)
		
		$counter = $source.Count;
		$i = 0;
		       
		if($counter -eq "1")
		{
		    $target.AutoConnect($shape[0], $Down, $actionConnector)
		}
		else
		{
		    for($i; $i -lt $counter; $i++)
		    {
		        if($i -eq "0")
		        {
		          	$target.AutoConnect($shape[$i], $Down, $actionConnector)
		        }
		        else
		        {
		            $shape[$i-1].AutoConnect($shape[$i], $Down, $actionConnector)
		        } 
		    }
		}
		$i = 0;
	}

	function MakeConnectors
	{
		Param($source, $target)
		
		$conn = $visPage.Drop($connector, 0, 0)
	    $start = $conn.CellsU("BeginX").GlueTo($source.CellsU("PinX"))
	    $end = $conn.CellsU("EndX").GlueTo($target.CellsU("PinX"))		
	}

	function DrawShapes
	{
		Param($shape, $startLine, $startRow, $visResizeDirE, $visResizeDirW, $visResizeDirS)
		
		$drawnShape = $visPage.drop($shape, $startLine, $startRow)
		
		$drawnShape.resize("visResizeDirE", $visResizeDirE, "visCentimeters")
		$drawnShape.resize("visResizeDirS", $visResizeDirS, "visCentimeters")
		$drawnShape.resize("visResizeDirW", $visResizeDirW, "visCentimeters")

		return $drawnShape			
	}

	function DrawWFProperties
	{
	    Param($count, $propertyCounter, $wfPropertiesShapes)
        $wfPropertiesShapes = @()

	    #Drawing the Workflowproperties
        #get them by the previous generated Hashtable
        #key is the name of the workflow
		
		#Are there any Workflows present?
	    if($wfHT[$wf[$count]].length -gt 0)
	    {
	        foreach($value in $wfHT[$wf[$count]].Split('#'))
	        {
	            $wfPropertiesShapes += $visPage.drop($wfActivityShape, $startLine, $startRow)
	            Set-VisioText $value
	            $propertyCounter++; 
	        }
            
            #arrange WF-Properties      
	        $j = 0;
	        for($j; $j -lt $propertyCounter; $j++)
	        {
                #is there more than one Property?
	            if($j -eq "0") 
	            {
	                $currentWorkflowShape.AutoConnect($wfPropertiesShapes[$j], $Right, $connector)
	            }
	            elseif(!($propertyCounter -eq "1"))
	            {
	                $wfPropertiesShapes[$j-1].AutoConnect($wfPropertiesShapes[$j], $Down, $connector)
	            }
	        }
	    }
        else
        {
            $propertyCounter = 0
        }
	    
        return $wfPropertiesShapes
	}

	function SelectAndGroupShapesFromArray
	{
		#First deselect all current selected shapes
		#then select and group all shapes that are stored 
		#in $allShapes
        Select-VisioShape none
        $allShapes = @();
		if($args)
        {
            foreach($arg in $args)
            {
                if($arg -ne $null)
                {
                    $allShapes += $arg

                }
            }
        }

		Select-VisioShape $allShapes
        $group = New-VisioGroup
        return $group
	}

    function DrawRectangleBehindWF
    {
        #get height, width, pinX and pinY from the group
        #you get the values always from the selected items 
        #it´s a little bit weird but you have to select the corresponding
        #attribute of the selected attribute 
		#e.g. $height = $height.height

        $height = Get-VisioShapeCell -Height
        $height = $height.height
        $width = Get-VisioShapeCell -Width
        $width = $width.width

        $pinX = Get-VisioShapeCell -PinX
        $pinX = $pinX.pinX
        $pinY = Get-VisioShapeCell -PinY
        $pinY = $pinY.pinY

        #Draw the Rectangle behind the WF and WF-Properties
        $rectangle = DrawShapes $rectShape3
        #set the new values
        Set-VisioShapeCell -Height $height -Width "90mm" -PinX "145.75mm" -piny $piny
        $rectangle.sendtoBack()
        #resize the rectangle
        $rectangle.Resize("visResizeDirN", 0.5, "visCentimeters")
        $rectangle.Resize("visResizeDirS", 0.5, "visCentimeters")

        #get the latest position for the next workflow + 0,5 centimeters
        $newPinY = Get-VisioShapeCell -PinY
        $newPinY = $newPinY.pinY
        $newPinY = $newPinY.replace("mm","")
        $newPinY = [double]$newPinY
        $newPinY = $newPinY*2
        $newPinY = $newPinY-0.5
        $newPinY = [string]$newPinY+"mm"
        
        return $newPiny

    }

    function DrawRectangleBehindAll
    {
        $height = Get-VisioShapeCell -Height
        $height = $height.height
        $width = Get-VisioShapeCell -Width
        $width = $width.width

        $pinX = Get-VisioShapeCell -PinX
        $pinX = $pinX.pinX
        $pinY = Get-VisioShapeCell -PinY
        $pinY = $pinY.pinY

        $rectangle = DrawShapes $rectShape4
        Set-VisioShapeCell -Height $height -Width $width -PinX $pinX -piny $piny
        $rectangle.sendtoBack()
        #resize the rectangle
        $rectangle.Resize("visResizeDirNW", 0.5, "visCentimeters")
        $rectangle.Resize("visResizeDirSE", 0.5, "visCentimeters")
        return $rectangle
    }
    
	

#endregion

function GenerateVisio
{
    #load visio-snappin
	Import-Module VisioPS
	
	$visPage = New-VisioPage -Name $MPRName
	
    #set page-rotation to landscape
    Set-VisioPageLayout –Orientation Landscape 
    
	#region drop mpr, set and wf shapes
	
    #if Set-Transition then...
    if($policyRuleType -eq "SetTransition")
    {
        $requestor = "Set-Transition"
    }
    
    #if Request then...
	
        #MPR
        $MPRRoot = DrawShapes -shape $mprShape -startLine $startLine -startRow $startRow -visResizeDirE 0.75 -visResizeDirW 0.75
    	Set-VisioText $MPRName
	
        #Requestor
        $requestorRoot = DrawShapes -shape $setShape -startLine ($startRow-(2.5)) -startRow ($startLine-(1.2)) -visResizeDirE 0.75 -visResizeDirW 0.75

        if($requestorFilter -ne $null)
        {
            $requestorRoot.text = $requestor + "`n - `n" + $requestorFilter
        }
        elseif($requstorMembers -ne $null)
        {
            $requestorRoot.text = $requestor + "`n - `n Explicit Member"
        }
        else
        {
            $requestorRoot.text = $requestor
        }

        #Array for ActionType Shapes
        $actionTypeShapes = @()
        #generate ActionType Shapes and fill the Array
        foreach($actionType in $actionTypes)
        {
            $actionTypeShapes += $visPage.Drop($mprActionShape, $startRow, $startLine)
            
            Set-VisioText $actionType
        }

        #Array for ActionParameter Shapes
        $actionParameterShapes = @()
        foreach($actionParameter in $actionParameters)
        {
            $actionParameterShapes += $visPage.Drop($setParameterShape, $startRow, $startLine)
            
            Set-VisioText $actionParameter
        }
        
        if($targetBefore -ne $null)
        {
            if($targetAfter -ne $null)
            {
                #targetBefore
                $targetBeforeShape = DrawShapes -shape $setShape -startline $startLine -startrow ($startRow-(1.2)) -visResizeDirE 0.75 -visResizeDirW 0.75
                if($targetBeforeFilter -ne $null)
                {
                    $targetBeforeShape.text = $targetBefore + "`n - `n" + $targetBeforeFilter
                }
                elseif($targetBeforeMembers -ne $null)
                {
                    $targetBeforeShape.text = $targetBefore + "`n - `n Explicit Member"
                }
                else
                {
                    $targetBeforeShape.text = $targetBefore
                }
            }
            else
            {
                #targetAfter becomes targetBefore
                $targetAfter = $targetBefore;
                $targetBefore = $null;
            }
        }
        if($targetBefore -eq $null)
        {
            #draw a small circle to make connections easier 
            $circleShape = Get-VisioMaster "Process (circle)" $floshp_m
            $circle = DrawShapes -shape $circleShape -startline 0.4735 -startrow (-1.1) -visResizeDirE (-2.4) -visResizeDirS (-2.4)          
        }
        
        #target after
        $targetAfterShape = DrawShapes -shape $setShape -startline ($startLine+(2.5)) -startrow ($startRow-(1.2)) -visResizeDirE 0.75 -visResizeDirW 0.75

        if($targetAfterFilter -ne $null)
        {
            $targetAfterShape.text = $targetAfter + "`n - `n" + $targetAfterFilter
        }
        elseif($targetAfterMembers -ne $null)
        {
            $targetAfterShape.text = $targetAfter + "`n - `n Explicit Member"
        }
        else
        {
            $targetAfterShape.text = $targetAfter
        }
        
        #endregion       
	
       #######################################################################################
       ########################        Arrange drawn Visio-Shapes         ####################
       #######################################################################################
       
        if($targetAfter -ne $null)
        {
            PinToRequestor -source $actionTypes -target $requestorRoot -shape $ActionTypeShapes
            $i = 0;
            if($circle -eq $null)
            {
                #connect targetBefore and targetAfter to the MPRRoot
                $MPRRoot.AutoConnect($targetBeforeShape, $Auto, $connector)
				MakeConnectors -source $requestorRoot -target $targetAfterShape

            }
            else
            {
                MakeConnectors -source $circle -target $requestorRoot
				MakeConnectors -source $circle -target $targetaftershape
				MakeConnectors -source $circle -target $mprroot
			}
        }
        else
        {
            $MPRRoot.AutoConnect($targetAfter, $Down, $connector)

            PinToRequestor -source $actionTypes -target $requestorRoot -shape $ActionTypeShapes
    
            #connect targetBefore and targetAfter to the MPRRoot
			MakeConnectors -source $requestorRoot -target $targetBeforeShape
        }
	
        #Draw a rectangle behind the sets
		$rect1 = DrawShapes -shape $rectShape1 -startLine ($startLine-(2)) -startRow ($startRow-(1.17)) -visResizeDirE 10.15 -visResizeDirS 0.16 
        $rect1.SendToBack()
		
		$actionParameterCount = $actionParameters.Count
		
		#pin ActionParameter to the Target
		#ActionParameter anhängen
        if(!($targetAfter -eq $null))
        {
			#pin TargetBefore to TargetAfter
			PinToRequestor -source $actionParameters -target $targetAfterShape -shape $actionParameterShapes    
        }
        else
        {
			PinToRequestor -source $actionParameters -target $targetBeforeShape -shape $actionParameterShapes
        }

        ############################################################################################
        ##############       Here is the beginning for drawing the workflows          ##############
        ############################################################################################

        #count how many workflows are there
        $workflowCount = $wf.Count;  
        
        #WorkflowPropertie Array
        $wfPropertiesShapes = @();
        $wfShapeArray = @();
        $wfLinePosition = 0;
        $propertyCounter = 0;
        $workflowPinX
        $newPinY
        #$globalPropertyCounter = 0;
        
        for($i; $i -lt $workflowCount; $i++)
		{
		    if($workflowCount -eq "1")
		    {
			    $currentWorkflowShape = DrawShapes -shape $wfShape -startLine ($startLine+(5.0)) -startRow ($startRow-(1.2)) -visResizeDirE 0.75 -visResizeDirW 0.75
				Set-VisioText $wf[0].Replace("Type:","`n`nType:");
					
				MakeConnectors -source $targetAfterShape -target $currentWorkflowShape  
		        $propertyShapes = DrawWFProperties -count $i -propertyCounter $propertyCounter -wfPropertiesShapes $wfPropertiesShapes
                
                SelectAndGroupShapesFromArray $currentWorkflowShape $propertyShapes
                DrawRectangleBehindWF

		    }
		    elseif($workflowCount -gt "1")
        	{
                if($i -eq 0)
                {
			        $currentWorkflowShape = DrawShapes -shape $wfShape -startLine ($startLine+(5.0)) -startRow (($startRow-(1.2)+($propertyCounter*(-1)))) -visResizeDirE 0.75 -visResizeDirW 0.75
                    #get current X-coordinate for later arrangement
                    $workflowPinX = Get-VisioShapeCell -pinx
                    $workflowPinX = $workflowPinX.PinX
                }
                else
                {
                    $currentWorkflowShape = DrawShapes -shape $wfShape -visResizeDirE 0.75 -visResizeDirW 0.75   
                    Set-VisioShapeCell -PinX $workFlowPinX -PinY $newPinY 
                }
                
               
				Set-VisioText $wf[$i].Replace("Type:","`n`nType:");
					
                if($i -eq "0")
                {
                    MakeConnectors -source $targetAfterShape -target $currentWorkflowShape
                }
                else
                {
                    MakeConnectors -source $wfShapeArray[$i-1] -target $currentWorkflowShape
                }
					
                $wfShapeArray += $currentWorkflowShape; 
                #$propertyCounter = 0; 
		        $propertyShapes = DrawWFProperties -count $i -propertyCounter $propertyCounter
                
                if($propertyShapes -ne $null)
                {
	                SelectAndGroupShapesFromArray $currentWorkflowShape $propertyShapes            
                }
                else
                {
                    SelectAndGroupShapesFromArray $currentWorkflowShape
                }
                $newPinY = DrawRectangleBehindWF
                
            }
        }

	    
    #Draw the big rectangle behind the whole arrangement
    Select-VisioShape none
    Select-VisioShape all
    $group = new-visiogroup


	$rect4 = DrawRectangleBehindAll

    $rect4.sendToBack()

    $legend = $visPage.drop($legendShape, 0, 0)
    $rect4.Autoconnect($legend, $right, $legendConnector)

    $visPage.ResizeToFitContents()
    $vispage.CenterDrawing()
    $i = 0;
}

#End Functions


[System.Xml.XmlDocument] $xd = new-object System.Xml.XmlDocument

$file = resolve-path($filePath)
$xd.load($file)

foreach($MPR in $xd.ManagementPolicyRules.MPR)
{
	foreach($node in $MPR){
        
        #what kind of policy is this MPR? Set Transition or Request
        $policyRuleType = $node.ManagementPolicyRuleType.InnerText
        #Name of the MPR
		$MPRName = $node.Name.Replace(":", "_");
        #find Requestor-Set in the xml
        $requestor = $xd.ManagementPolicyRules.SETs.SET| Where-Object{$_.ID -eq $node.PrincipalSet}
        
        #if Filter is present then save it in $requestorFilter
        $requestorFilter = $requestor.Filter.Filter.InnerXml
        #if manually managed membership, then save the ExplicitMember value in $requestorMembers
        $requstorMembers = $requestor.ExplicitMember.Value
        #Name of the Requestor-Set
        $requestor = $requestor.Name

        #find affected parameters
        $actionParameters = @();
        foreach($parameter in $node.ActionParameter)
        {
            if($parameter.Value -eq "*")
            {
                $parameter = "All Attributes"
            }
            else
            {
                $actionParameters += $parameter.Value
            }
        }
		
        #find actions
        $actionTypes = @();       
        foreach($type in $node.ActionType.Value)
        {
            if($type -eq "TransitionIn")
            {
                $type = "Transition In"
            }
            if($type -eq "TransitionOut")
            {
                $type = "Transition Out"
            }
            $actionTypes += $type;
        }

        #Set TargetBefore
        $targetBefore = $xd.ManagementPolicyRules.SETs.Set | Where-Object{$_.ID -eq $node.ResourceCurrentSet}
        #if Filter is present then save it in  $targetBeforeFilter
        $targetBeforeFilter = $targetBefore.Filter.Filter.InnerXml
        #if manually managed membership, then save the ExplicitMember value in $targetBeforeMembers
        $targetBeforeMembers = $targetBefore.ExplicitMember.Value
        $targetBefore = $targetBefore.Name
        
        #Set TargetAfter
        $targetAfter = $xd.ManagementPolicyRules.SETs.Set | Where-Object{$_.ID -eq $node.ResourceFinalSet}	
		#if Filter is present then save it in $targetAfterFilter
		$targetAfterFilter = $targetAfter.Filter.Filter.InnerXml
        #if manually managed membership, then save the ExplicitMember value in $targetAfterMembers
        $targetAfterMembers = $targetAfter.ExplicitMember.Value
        $targetAfter = $targetAfter.Name
        
        #now get the name and the properties of each workflow
        #store each wf-name in a array so you can later access
        #the hashtable with it´s name
        $wf = @();

        #first make a hashtable for each wf
        #and store the name and properties in it.
        #later you can access this ht by simply put
        #the name of the wf as key und you get back 
        #the values you stored in.
        #values are separated by hashtags (#)
        #you can split these by using the split()-operator
		
        $wfHT = @{}
        $wfProperties = @()
        $tempWF = @()
        $tempWF += $node.ChildNodes | Where-Object{$_.Type -eq "WorkflowDefinition"}

        foreach($workFlow in $tempWF)
        {
            $tempWFValue = $workflow.Value
            $tempWholeWF = $xd.ManagementPolicyRules.Workflows.Workflow | Where-Object{$_.ID -eq $tempWFValue}
			$tempWFType = "Type: "+$tempWholeWF.requestphase.innertext;
            $tempWFName = $tempWholeWF.Name + $tempWFType;

            $wf += $tempWFName;

            $wfProperties = ""

            foreach($wfProperty in $tempWholeWF.XOML.SequentialWorkflow.ChildNodes)
            {
                if($wfProperty.Name -like "*Activity*")
                {
                    $wfProperties += $wfProperty.LocalName+"#"
                }
            }
            if($wfProperties.length -gt 0)
            {
                $wfProperties = $wfProperties.TrimEnd("#")  
            }

            $wfHT.Add($tempWFName, $wfProperties);
        } 
        Clear-Variable tempWF
	}

    GenerateVisio $policyRuleType, $MPRName, $requestor, $actionParameters, $actionTypes, $requestorFilter, $requstorMembers, $targetBefore, $targetBeforeFilter, $targetBeforeMembers, $targetAfter, $targetAfterFilter, $targetAfterMembers, $wf, $wfHT
}

#Save the Visio document and export the pages as PNG
#automatically create a folder with the current date

cd $savePath
$dirName = Get-date -Format d
mkdir $dirName
cd $dirName
$workDir = pwd
$p1 = Get-VisioPage "Page-1"
Set-VisioPage $p1
Remove-VisioPage
$visDocSaveName = "$workDir\$dirName.vsd"
$visDoc.SaveAs($visDocSaveName)

Export-VisioPage "$($workDir)\.png" -AllPages 