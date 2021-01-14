$Imported_functions={

Function Format-HtmlTable {

<#
.SYNOPSIS
Define the style and color of tables

.DESCRIPTION
Define the style and color of tables

.PARAMETER Contents
The objects to convert into html table

.PARAMETER Title
Define the table head

.PARAMETER ColorItems
It will color the table if this switch enabled.

.EXAMPLE
Format-HtmlTable -Contents <objects> -Title <tableTitleString> -ColorItems

#>
  
    Param (
        [Parameter(Mandatory=$true,
                    Position=0)]
        [Object[]] $Contents,
        [String] $Title,
        [Int] $TableBorder=1,
        [Int] $Cellpadding=0,
        [Int] $Cellspacing=0,
        [Parameter(Mandatory=$true,
                    ParameterSetName="Item")]
        [Switch] $ColorItems,
        [Parameter(ParameterSetName="Item")]
        [String] $DeletedBgcolor="#FFFF00",
        [Parameter(ParameterSetName="Item")]
        [String] $NewBgcolor="#00FF00",
        [Parameter(Mandatory=$true,
                    ParameterSetName="Property")]
        [Switch] $ColorPorperties

    )
    $TableHeader = "<table border=$TableBorder cellpadding=$Cellpadding cellspacing=$Cellspacing"

    # Add title
    $HTMLContent = $Contents | ConvertTo-Html -Fragment -PreContent "<H1>$Title</H1>"
    #Add table border
    $HTMLContent = $HTMLContent -replace "<table", $TableHeader
    # Strong table header
    $HTMLContent = $HTMLContent -replace "<th>", "<th><strong>"
    $HTMLContent = $HTMLContent -replace "</th>", "</strong></th>"

    # Color Items. Yellow highlight deleted items and Green highlight new items.
    If ($ColorItems) {
        
        for ($i=0;$i -lt $HTMLContent.count;$i++) {
            # Color Deleted items.
            If ($HTMLContent[$i].IndexOf("<td>Deleted</td>") -ge 0) {
                $InsertString = " bgcolor=" + $DeletedBgcolor
                $FindStr = "<tr"
                $Position = $HTMLContent[$i].IndexOf($FindStr) + $FindStr.Length
                $HTMLContent[$i] = $HTMLContent[$i].Insert($Position,$InsertString)
            }

            # Color New items.
            If ($HTMLContent[$i].IndexOf("<td>New</td>") -ge 0) {
                $InsertString = " bgcolor=" + $NewBgcolor
                $FindStr = "<tr"
                $Position = $HTMLContent[$i].IndexOf($FindStr) + $FindStr.Length
                $HTMLContent[$i] = $HTMLContent[$i].Insert($Position,$InsertString)
            }

        }
    }

    # Color Properties.
    If ($ColorPorperties) {
        # Coming soon...
    }

    Return $HTMLContent
}

Function Compare-Objects {

<#
.SYNOPSIS
Compare 2 groups Objects to classifid the results by "Same","New","Deleted"

.DESCRIPTION
Compare 2 groups Objects to classifid the results by "Same","New","Deleted"

.PARAMETER ReferenceItems
The objects to be compared with

.PARAMETER CurrentItems
The objects to compare

.PARAMETER Referproperties
Sepcified the properties to compare. If it's "All", it means they are 2 different object if their have a similiar different property 

.EXAMPLE
Compare-Objects -ReferenceItems <ReferObjects> -CurrentItems <DiffObjects> -Referproperties "FlightId","Name"

#>

    Param (
        [Parameter(Mandatory=$true)]
        [Object[]] $ReferenceItems,
        [Parameter(Mandatory=$true)]
        [Object[]] $CurrentItems,
        [Parameter(Mandatory=$true)]
        [String[]] $Referproperties
    )

    Write-Debug "Start Flights' comparing ..."

    $Results = @()
    
    $Rlts = Compare-Object $ReferenceItems $CurrentItems -Property $Referproperties -IncludeEqual

    foreach ($Rlt in $Rlts) {
        

        If ($Rlt.SideIndicator -eq '==') {

            $EqualItems = $Rlts | ? Name -EQ $Rlt.Name | Select-Object $Referproperties 
            $EqualItems | Add-Member -NotePropertyName Record -NotePropertyValue Same
            $Results += $EqualItems
        }

        If ($Rlt.SideIndicator -eq '<=') {
            $DeletedItems = $Rlts | ? Name -EQ $Rlt.Name | Select-Object $Referproperties
            $DeletedItems | Add-Member -NotePropertyName Record -NotePropertyValue Deleted
            $Results += $DeletedItems
        }

        If ($Rlt.SideIndicator -eq '=>') {
            $NewItems = $Rlts | ? Name -EQ $Rlt.Name | Select-Object $Referproperties
            $NewItems | Add-Member -NotePropertyName Record -NotePropertyValue New
            $Results += $NewItems
        }
    }
    Write-Debug "Compare compeleted!"
    Return $Results
}

Function Compare-Properties {

<#
.SYNOPSIS
Compare 2 object its properties one by one.

.DESCRIPTION
Compare 2 object its properties one by one.

.PARAMETER referenceobject
The object to be compared with

.PARAMETER differenceobject
The object to compare

.PARAMETER SepciProperties
Sepcified the properties to compare. If it's "All", it will foreach all properties, and compare the property name and value.

.EXAMPLE
Compare-Properties -referenceobject $Object1 -differenceobject $Object2 -SepciProperties 'FlightState','UpdatedBy','UpdatedTime'

#>

    [CmdletBinding()]
    Param(
        # reference object
        [ValidateNotNull()]
        $referenceobject,

        # Compared object
        [ValidateNotNull()] 
        $differenceobject,  
        
        #specific properties to compare 
        [ValidateNotNull()] 
        [String[]]$SepciProperties="All"    
            
        )

    # List out all properties of comparing, compared objects
    $RefPropList = $referenceobject | Get-Member |Where-Object membertype -eq NoteProperty | select-object -ExpandProperty name
    $DifPropList = $differenceobject | Get-Member |Where-Object membertype -eq NoteProperty | select-object -ExpandProperty name
    
    # Compare properties' Name
    $propsCompare = Compare-Object $RefPropList $DifPropList -IncludeEqual
    $difPropsOnly = ($propsCompare | Where-Object sideindicator -EQ "=>").inputobject
    $refPropsOnly = ($propsCompare | Where-Object sideindicator -EQ "<=").inputobject

    if($SepciProperties -eq "All" ){$identicalprops = ($propscompare | Where-Object sideindicator -EQ "==").inputobject}
    else{$identicalprops=$SepciProperties}
    

    [string]$result= @()

    if ($identicalprops) {

       foreach ($identicalprop in $identicalprops){
    
          $RefPropValue = $referenceobject.$identicalprop
          if ([string]::IsNullOrEmpty($RefPropValue)){$RefPropValue = "null"}

          $DifPropValue = $differenceObject.$identicalprop
          if ([string]::IsNullOrEmpty($DifPropValue)){$DifPropValue = "null"}
          
          if ($RefPropValue -ne $DifPropValue){

              $PropValueChangedString =  "`n Property  Changed:Value of $identicalprop Change from $RefPropValue to $DifPropValue"
              $result += $PropValueChangedString
          }
       }
    }
    if ($RefPropsOnly) {

        $RefPropsOnlyString =  -join "$RefPropsOnly",","
        $RefPropsOnlyString =  -join "`n Properties Deleted:","$RefPropsOnlyString"
        $result += $RefPropsOnlyString
    }
    if ($DifPropsOnly) {
        
        $DifPropsOnlyString = -join "$DifPropsOnly","," 
        $DifPropsOnlyString =  -join "`n Properties Added:","$DifPropsOnlyString"
        $result += $DifPropsOnlyString
    } 

    return $result
}

Function Compare-Flights {

    Param (  

       [Object[]]$ReferObjects,
       [Object[]]$DiffObjects,
       [Switch]$CompareProperty
    )
  
    $Rst = Compare-Objects -ReferenceItems $ReferObjects -CurrentItems $DiffObjects -Referproperties "FlightId","Name"

    If ($Rst) {
             
        $Rst = $Rst| Select-Object FlightId,Name,Record
        
        if($CompareProperty){

        $Same=($Rst|?{$_.Record -eq "Same"}).FlightId

        foreach($s in $Same){
    
           $Object1=$ReferObjects|?{$_.FlightId -eq $s}|Select-Object FlightId,FlightState,UpdatedBy,UpdatedTime
           $Object2=$DiffObjects|?{$_.FlightId -eq $s}|Select-Object FlightId,FlightState,UpdatedBy,UpdatedTime

           $PropertyRst=Compare-Properties -referenceobject $Object1 -differenceobject $Object2 -SepciProperties 'FlightState','UpdatedBy','UpdatedTime'
           If(!$PropertyRst){continue}
           Else{
             $Rst|?{$_.FlgihtId -eq $s} | %{ $_.Record = $PropertyRst}
           }
        }
      }
    }
        
    Write-Host "Comparing For Flights completed!" -ForegroundColor Green
    Return $Rst
}

Function Compare-CenteralFlights {
    
    Param (  

       [Object[]]$ReferObjects,
       [Object[]]$DiffObjects,
       [Switch]$CompareProperty
    )
  
    $Rst = Compare-Objects -ReferenceItems $ReferObjects -CurrentItems $DiffObjects -Referproperties "Id","Name"

    If ($Rst) {
             
        $Rst = $Rst| Select-Object Id,Name,Record
        
        if($CompareProperty){

        $Same=($Rst|?{$_.Record -eq "Same"}).Id

          foreach($s in $Same){
          
             $Object1=$ReferObjects|?{$_.Id -eq $s}|Select-Object ID,FlightVersion,GlobalVersion,LastModifiedTime,PermanentlyDeleted
             $Object2=$DiffObjects|?{$_.Id -eq $s}|Select-Object ID,FlightVersion,GlobalVersion,LastModifiedTime,PermanentlyDeleted
          
             $PropertyRst = Compare-Properties -referenceobject $Object1 -differenceobject $Object2 -SepciProperties 'FlightVersion','GlobalVersion','LastModifiedTime','PermanentlyDeleted'

             If(!$PropertyRst){continue}

             Else{
               
               $Rst|?{$_.Id -eq $s} | %{ $_.Record = $PropertyRst}}
           }
         }
        
    Write-Host "Comparing For CenteralFlights completed!" -ForegroundColor Green
    Return $Rst
    }
}

}