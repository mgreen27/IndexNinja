<# Powershell - Tested on version 4 
.SYNOPSIS
    Gets filenames which have been indexed by Windows desktop search and have contents matching filter keywords
        
    Name: Get-IndexNinja.ps1
    Version: 1.1
    Author: @mgreen27

.DESCRIPTION
    Script to search for keywords in Windows Index and output lists for each keyword.
    Add in variables in param below - Search path, Results Path and Key Words (text file)
    Please ensure all expected items are indexed
    Install all appropriate windows filters - in my initial usecase we needed to install a Office 2010 filter to idex xlsx and docx files which are not indexed by default and utilise a binary encoding.

    Get-IndexItem from Technet - https://gallery.technet.microsoft.com/scriptcenter/Get-IndexedItem-PowerShell-5bca2dae
    Office 2010 Filter update - https://www.microsoft.com/en-us/download/details.aspx?id=17062
.EXAMPLE
    Get-IndexNinja.ps1
    Get-IndexNinja.ps1 -SearchPath $SearchPath -ResultsPath $ResultsPath -list 
#>
 param (
    [string]$SearchPath = (Get-Item -Path ".\" -Verbose).FullName, # defaults to script path
    [string]$ResultsPath = (Get-Item -Path ".\" -Verbose).FullName, # defaults to script path
    [string]$Keywords = "Test,Test Phrase, test4", # Add comman seperated values here if not using keyword txt file
    [string]$List = $Null # Default uses Keywords above, enter list for listfile
 )

Function Get-IndexedItem {
[CmdletBinding()]
Param ( [Alias("Where","Include")][String[]]$Filter , 
        [String]$path, 
        [Alias("Sort")][String[]]$orderby, 
        [Alias("Top")][int]$First,
        [Alias("Group")][String]$Value, 
        [Alias("Select")][String[]]$Property, 
        [Switch]$recurse,
        [Switch]$list,
        [Switch]$NoFiles)

#Alias definitions take the form  AliasName = "Full.Cannonical.Name" ; 
#Any defined here will be accepted as input field names in -filter and -OrderBy parameters
#and will be added to output objects as AliasProperties. 
 $PropertyAliases   = @{Width         ="System.Image.HorizontalSize"; Height        = "System.Image.VerticalSize";  Name    = "System.FileName" ; 
                        Extension     ="System.FileExtension"       ; CreationTime  = "System.DateCreated"       ;  Length  = "System.Size" ; 
                        LastWriteTime ="System.DateModified"        ; Keyword       = "System.Keywords"          ;  Tag     = "System.Keywords"
                        CameraMaker  = "System.Photo.Cameramanufacturer"}

 $fieldTypes = "System","Photo","Image","Music","Media","RecordedTv","Search","Audio" 
#For each of the field types listed above, define a prefix & a list of fields, formatted as "Bare_fieldName1|Bare_fieldName2|Bare_fieldName3"
#Anything which appears in FieldTypes must have a prefix and fields definition. 
#Any definitions which don't appear in fields types will be ignored 
#See http://msdn.microsoft.com/en-us/library/dd561977(v=VS.85).aspx for property info.  
 
 $SystemPrefix     = "System."            ;     $SystemFields = "ItemName|ItemUrl|FileExtension|FileName|FileAttributes|FileOwner|ItemType|ItemTypeText|KindText|Kind|MIMEType|Size|DateModified|DateAccessed|DateImported|DateAcquired|DateCreated|Author|Company|Copyright|Subject|Title|Keywords|Comment|SoftwareUsed|Rating|RatingText"
 $PhotoPrefix      = "System.Photo."      ;      $PhotoFields = "fNumber|ExposureTime|FocalLength|IsoSpeed|PeopleNames|DateTaken|Cameramodel|Cameramanufacturer|orientation"
 $ImagePrefix      = "System.Image."      ;      $ImageFields = "Dimensions|HorizontalSize|VerticalSize"
 $MusicPrefix      = "System.Music."      ;      $MusicFields = "AlbumArtist|AlbumID|AlbumTitle|Artist|BeatsPerMinute|Composer|Conductor|DisplayArtist|Genre|PartOfSet|TrackNumber"
 $AudioPrefix      = "System.Audio."      ;      $AudioFields = "ChannelCount|EncodingBitrate|PeakValue|SampleRate|SampleSize"
 $MediaPrefix      = "System.Media."      ;      $MediaFields = "Duration|Year"
 $RecordedTVPrefix = "System.RecordedTV." ; $RecordedTVFields = "ChannelNumber|EpisodeName|OriginalBroadcastDate|ProgramDescription|RecordingTime|StationName"
 $SearchPrefix     = "System.Search."     ;     $SearchFields = "AutoSummary|HitCount|Rank|Store"
 
 if ($list)  {  #Output a list of the fields and aliases we currently support. 
    $( foreach ($type in $fieldTypes) { 
          (get-variable "$($type)Fields").value -split "\|" | select-object @{n="FullName" ;e={(get-variable "$($type)prefix").value+$_}},
                                                                            @{n="ShortName";e={$_}}    
       }
    ) + ($PropertyAliases.keys | Select-Object  @{name="FullName" ;expression={$PropertyAliases[$_]}},
                                                @{name="ShortName";expression={$_}}
    ) | Sort-Object -Property @{e={$_.FullName -split "\.\w+$"}},"FullName" 
  return
 }  
  
#Make a giant SELECT clause from the field lists; replace "|" with ", " - field prefixes will be inserted later.
#There is an extra comma to ensure the last field name is recognized and gets a prefix. This is tidied up later
 if ($first)    {$SQL =  "SELECT TOP $first "}
 else           {$SQL =  "SELECT "}
 if ($property) {$SQL += ($property -join ", ") + ", "}
 else {
    foreach ($type in $fieldTypes) { 
        $SQL += ((get-variable "$($type)Fields").value -replace "\|",", " ) + ", " 
    }
 }   
  
#IF a UNC name was specified as the path, build the FROM ... WHERE clause to include the computer name.
 if ($path -match "\\\\([^\\]+)\\.") {
       $sql += " FROM $($matches[1]).SYSTEMINDEX WHERE "  
 } 
 else {$sql += " FROM SYSTEMINDEX WHERE "} 
 
#If a WHERE condidtion was provided via -Filter, add it now   

 if ($Filter) { #Convert * to % 
                $Filter = $Filter -replace "(?<=\w)\*","%"
                #Insert quotes where needed any condition specified as "keywords=stingray" is turned into "Keywords = 'stingray' "
                $Filter = $Filter -replace "\s*(=|<|>|like)\s*([^\''\d][^\d\s\'']*)$"  , ' $1 ''$2'' '
                # Convert "= 'wildcard'" to "LIKE 'wildcard'" 
                $Filter = $Filter -replace "\s*=\s*(?='.+%'\s*$)" ," LIKE " 
                #If a no predicate was specified, use the term in a contains search over all fields.
                $filter = ($filter | ForEach-Object {
                                if ($_ -match "'|=|<|>|like|contains|freetext") {$_}
                                else {"Contains(*,'$_')"}
                }) 
                #if $filter is an array of single conditions join them together with AND 
                  $SQL += $Filter -join " AND "  } 
                  
 #If a path was given add SCOPE or DIRECTORY to WHERE depending on whether -recurse was specified. 
 if ($path)     {if ($path -notmatch "\w{4}:") {$path = "file:" + (resolve-path -path $path).providerPath}  # Path has to be in the form "file:C:/users" 
                $path  = $path -replace "\\","/"
                if ($sql -notmatch "WHERE\s$") {$sql += " AND " }                       #If the SQL statement doesn't end with "WHERE", add "AND"  
                if ($recurse)                  {$sql += " SCOPE = '$path' "       }     #INDEX uses SCOPE <folder> for recursive search, 
                else                           {$sql += " DIRECTORY = '$path' "   }     # and DIRECTORY <folder> for non-recursive
 }   
 
 if ($Value) {
                if ($sql -notmatch "WHERE\s$") {$sql += " AND " }                       #If the SQL statement doesn't end with "WHERE", add "AND"  
                                                $sql += " $Value Like '%'" 
                                                $sql =  $SQL -replace "^SELECT.*?FROM","SELECT $Value, FROM"
 }
 
 #If the SQL statement Still ends with "WHERE" we'd return everything in the index. Bail out instead  
 if ($sql -match "WHERE\s*$")  { Write-warning "You need to specify either a path , or a filter." ; return} 
 
 #Add any order-by condition(s). Note there is an extra trailing comma to ensure field names are recognised when prefixes are inserted . 
 if ($Value) {$SQL =  "GROUP ON $Value, OVER ( $SQL )"}
 elseif ($orderby)  {$sql += " ORDER BY " + ($orderby   -join " , " ) + ","}             
 
 # For each entry in the PROPERTYALIASES Hash table look for the KEY part being used as a field name
 # and replace it with the associated value. The operation becomes
 # $SQL  -replace "(?<=\s)CreationTime(?=\s*(=|\>|\<|,|Like))","System.DateCreated" 
 # This translates to "Look for 'CreationTime' preceeded by a space and followed by ( optionally ) some spaces, and then
 # any of '=', '>' , '<', ',' or 'Like' (Looking for these prevents matching if the word is a search term, rather than a field name)
 # If you find it, replace it with "System.DateCreated" 
 
 $PropertyAliases.Keys | ForEach-Object { $sql= $SQL -replace "(?<=\s)$($_)(?=\s*(=|>|<|,|Like))",$PropertyAliases.$_}      

 # Now a similar process for all the field prefixes: this time the regular expression becomes for example,
 # $SQL -replace "(?<!\s)(?=(Dimensions|HorizontalSize|VerticalSize))","System.Image." 
 # This translates to: "Look for a place which is preceeded by space and  followed by 'Dimensions' or 'HorizontalSize'
 # just select the place (unlike aliases, don't select the fieldname here) and put the prefix at that point.  
 foreach ($type in $fieldTypes) { 
    $fields = (get-variable "$($type)Fields").value 
    $prefix = (get-variable "$($type)Prefix").value 
    $sql = $sql -replace "(?<=\s)(?=($Fields)\s*(=|>|<|,|Like))" , $Prefix
 }
 
 # Some commas were  put in just to ensure all the field names were found but need to be removed or the SQL won't run
 $sql = $sql -replace "\s*,\s*FROM\s+" , " FROM " 
 $sql = $sql -replace "\s*,\s*OVER\s+" , " OVER " 
 $sql = $sql -replace "\s*,\s*$"       , "" 
 
 #Finally we get to run the query: result comes back in a dataSet with 1 or more Datatables. Process each dataRow in the first (only) table
 write-debug $sql 
 $adapter = new-object system.data.oledb.oleDBDataadapter -argumentlist $sql, "Provider=Search.CollatorDSO;Extended Properties=’Application=Windows’;"
 $ds      = new-object system.data.dataset
 if ($adapter.Fill($ds)) { foreach ($row in $ds.Tables[0])  {
    #If the dataRow refers to a file output a file obj with extra properties, otherwise output a PSobject
    if ($Value) {$row | Select-Object -Property @{name=$Value; expression={$_.($ds.Tables[0].columns[0].columnname)}}}
    else {
        if (($row."System.ItemUrl" -match "^file:") -and (-not $NoFiles)) { 
               $obj = (Get-item -force -LiteralPath (($row."System.ItemUrl" -replace "^file:","") -replace "\/","\"))
               if (-not $obj) {$obj = New-Object psobject }
        }
        else { 
               if ($row."System.ItemUrl") {
                     $obj = New-Object psobject -Property @{Path = $row."System.ItemUrl"}
                     Add-Member -force -InputObject $obj -Name "ToString"  -MemberType "scriptmethod" -Value {$this.path} 
               }
               else {$obj = New-Object psobject }   
        }
        if ($obj) {
            #Add all the the non-null dbColumns removing the prefix from the property name. 
            foreach ($prop in (Get-Member -InputObject $row -MemberType property | where-object {$row."$($_.name)" -isnot [system.dbnull] })) {                            
                Add-member -ErrorAction "SilentlyContinue" -InputObject $obj -MemberType NoteProperty  -Name (($prop.name -split "\." )[-1]) -Value  $row."$($prop.name)"
            }                       
            #Add aliases 
            foreach ($prop in ($PropertyAliases.Keys | where-object {  ($row."$($propertyAliases.$_)" -isnot [system.dbnull] ) -and
                                                                       ($row."$($propertyAliases.$_)" -ne $null )})) {
                Add-member -ErrorAction "SilentlyContinue" -InputObject $obj -MemberType AliasProperty -Name $prop -Value ($propertyAliases.$prop  -split "\." )[-1] 
            }
            #Overwrite duration as a timespan not as 100ns ticks
            If ($obj.duration) { $obj.duration =([timespan]::FromMilliseconds($obj.Duration / 10000) )}
            $obj
        }
    }                               
 }}
}

## MAIN ##
$SearchPath = $SearchPath.Trimend('\')
$ResultsPath = $ResultsPath.Trimend('\')
if ($List) {$SearchItems = (Get-Content $List).trim()}
Else {$SearchItems = ($Keywords -split ',').trim()}

# Search file contents that have been indexed
ForEach ($Item in $SearchItems) {
    $Results = $Null
    $OutFile = "$ResultsPath\$Item.txt"
    If (Test-Path "$OutFile"){Remove-Item "$OutFile" -Force}
    $Item = $([Char]34)+$Item+$([Char]34)
    Try{
        "`nSearching $SearchPath using $Item..."
        $Results = Get-IndexedItem -filter $Item -path "$SearchPath" -recurse
        If ($Results) {
            $Results | ForEach-Object { $_.FullName} | Out-File -FilePath $OutFile -Encoding ascii
            "Results: $Item in $OutFile"
            }
        Else {"No Results: $Item"}
        }
    Catch{
        "Error: $Item"
        }
    } # End ForEach
 
"`nCompleted content search..."