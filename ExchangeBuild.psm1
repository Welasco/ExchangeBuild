<# 
 .Synopsis
  Get a list of all Exchange On-Premises Version, Release Date, Build Number.

 .Description
  Get a list of all Exchange On-Premise Version, ReleaseDate, Build Number.
  This Function reads all Exchange Servers On-Premise Version from Technet Website and export as an Powershell Object.

 .Parameter All
  List all Exchange On-Premises Version, Release Date, Build Number.

 .Parameter E4
  List all Exchange 4.0 Version, Release Date, Build Number.

 .Parameter E5
  List all Exchange 5.0 Version, Release Date, Build Number.

 .Parameter E55
  List all Exchange 5.5 Version, Release Date, Build Number.

 .Parameter E2000
  List all Exchange 2000 Version, Release Date, Build Number.

 .Parameter E2003
  List all Exchange 2003 Version, Release Date, Build Number.
 
 .Parameter E2007
  List all Exchange 2007 Version, Release Date, Build Number. 

 .Parameter E2010
  List all Exchange 2010 Version, Release Date, Build Number.

 .Parameter E2013
  List all Exchange 2013 Version, Release Date, Build Number.

 .Parameter E2016
  List all Exchange 2016 Version, Release Date, Build Number.    

 .Example
   # List all Exchange Server Versions:
   Get-ExchanbeBuild -All or Get-ExchanbeBuild

 .Example
   # List All Exchange 2013 Versions:
   Get-ExchanbeBuild -E2013

 .Example
   # List All Exchange 2016 Versions:
   Get-ExchanbeBuild -E2016

# A URL to the main website for this project.
ProjectUri = 'https://github.com/welasco/powershell'
#>

Function Get-ExchangeBuild{

    [cmdletbinding(DefaultParameterSetName="setA",PositionalBinding=$false)]
    [OutputType('System.Collections.ArrayList')]
     Param(
        [Parameter(ParameterSetName="setA")]
        [switch]$All,
        [Parameter(ParameterSetName="setB")]
        [switch]$E4,
        [Parameter(ParameterSetName="setC")]
        [switch]$E5,
        [Parameter(ParameterSetName="setD")]
        [switch]$E55,
        [Parameter(ParameterSetName="setE")]
        [switch]$E2000,
        [Parameter(ParameterSetName="setF")]
        [switch]$E2003,
        [Parameter(ParameterSetName="setG")]
        [switch]$E2007,
        [Parameter(ParameterSetName="setH")]
        [switch]$E2010,
        [Parameter(ParameterSetName="setI")]
        [switch]$E2013,
        [Parameter(ParameterSetName="setJ")]
        [switch]$E2016,
        [Parameter(ParameterSetName="setK")]
        [switch]$Cloud,
        [Parameter(ParameterSetName="setK")]
        [string]$Identity = ""

    )

    $url = 'https://technet.microsoft.com/library/hh135098.aspx'
    $content = Invoke-WebRequest -Uri $url
    $htmlcontent = $content.ParsedHtml
    $tablesections = $htmlcontent.getElementsByTagName('tbody')
    $tables = @()
    $tablesections | ForEach-Object {
        if ($_.innerText -like "Product name*"){
            $tables += $_.innertext
            }
                }
    $result = $tables -split '[\r\n]'

    $Months = "January","February","March","April","May","June","July","August","September","October","November","December"
    $TemplateObject = New-Object PSObject | Select-Object Version, ReleaseDate, BuildNumber
    $ReleaseDate = New-Object System.Collections.ArrayList
    $arrMulti = New-Object System.Collections.ArrayList
    $tempArr= New-Object System.Collections.ArrayList

    foreach($item in $result){
        if($item -ne "" -and $null -ne $item){
            if($item -like "Product name*"){
                $arrMulti.AddRange($tempArr)
                $tempArr.Clear()
            }
            else{
                foreach($Month in $Months){
                    if($item -like "*$Month*"){
                        $Version = $item -split $Month
                        $Version[1] = $Month+$Version[1]
                        $BuildNumber = $Version[1] -split '[ ]'
                        $ReleaseDate.AddRange($BuildNumber)
                        $BuildNumber = $BuildNumber[$BuildNumber.count-1].Substring(4)
                        $ReleaseDateYear = $ReleaseDate[$ReleaseDate.count-1].Substring(0,4)
                        $ReleaseDate.Removeat($ReleaseDate.Count-1)
                        foreach($ReleaseDateItem in $ReleaseDate){
                            $ReleaseDateResult += ($ReleaseDateItem+" ")
                        }
                        $ReleaseDateResult+=$ReleaseDateYear
                    }
                }
                $WorkingObject = $TemplateObject | Select-Object * 
                $WorkingObject.Version = $Version[0]
                $WorkingObject.ReleaseDate = $ReleaseDateResult
                $WorkingObject.BuildNumber = $Buildnumber
                $ReleaseDateResult = $null
                $ReleaseDate.Clear()
                $tempArr.Add($WorkingObject) | Out-Null
            }
        }
    }


    #$arrMulti.count

    if ($All) {$arrMulti}
    elseif ($E4) {$arrMulti | Where-Object{$_.Version -like "*4.0*"}}
    elseif ($E5) {$arrMulti | Where-Object{$_.Version -like "*5.0*"}}
    elseif ($E55) {$arrMulti | Where-Object{$_.Version -like "*5.5*"}}
    elseif ($E2000) {$arrMulti | Where-Object{$_.Version -like "*2000*"}}
    elseif ($E2003) {$arrMulti | Where-Object{$_.Version -like "*2003*"}}
    elseif ($E2007) {$arrMulti | Where-Object{$_.Version -like "*2007*"}}
    elseif ($E2010) {$arrMulti | Where-Object{$_.Version -like "*2010*"}}
    elseif ($E2013) {$arrMulti | Where-Object{$_.Version -like "*2013*"}}
    elseif ($E2016) {$arrMulti | Where-Object{$_.Version -like "*2016*"}}
    else{$arrMulti}
}

export-modulemember -function Get-ExchangeBuild