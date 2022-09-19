<#
 .Name
  Get Survey Result

 .Synopsis
  Script to generate Survey Result report.

 .Description
   It will lookup all survey data, extract data and show the result as a single row for each survey. It also read its JSON fields and break its data as column in excel.
   
 .Author
    Rinku Jain
 
#>
 

$surveyFormFolders = @("master:///sitecore/content/Shared Data/Survey Data Folder/Survey Form Folder")
    
$categoryCountProps = @{
    Parameters = @(
		@{ Name = "Language"; Title = "Select Language"; Source = "DataSource=/sitecore/system/languages" ; Editor = "droplist" }
    )
    Title = "Select Language"
	Description = "Select Language to generate report against"
    OkButtonName = "Proceed"
    CancelButtonName = "Abort"
	Width = 600
	Height = 400
}

$result = Read-Variable @categoryCountProps
$language = $Language.Name
if ($language -eq $null)
{
    $language = "*"
}
if ($result -ne "ok") {
    Exit
}
$StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch

Write-output "Evaluation  Started....";
Write-output $language;

$StopWatch.Start();



 $customList = New-Object System.Collections.Generic.List[PSObject]
try {
foreach ($folder in $surveyFormFolders) {
    get-childitem $folder -Language $language -recurse | where {$_.TemplateName -eq 'Survey Form Data'}| foreach {
    $esgitem =  $_
    $time = $esgitem."Time Statistics Json"
    $input = $esgitem."Survey Input Json" 
    $result = $esgitem."Survey Result Json"
    $timelist = $time |  ConvertFrom-Json
    $inputlist = $input |  ConvertFrom-Json
    $resultlist = $result |  ConvertFrom-Json
    $stdateactual=$esgitem.'Start Date'
    $compdateactual=$esgitem.'Completion Date'
      if (($stdateactual -ne $null ) -and (( get-date($stdateactual)  -Format "MM-dd-yyyy" ) -ne '01-01-0001')) {
          $stdate =   get-date($stdateactual)  -Format "MM-dd-yyyy" 
          $sttime = get-date($stdateactual) -Format "hh:mm:ss"
      }
      else{
           $stdate=  $null
            $sttime = $null
      }
       if (($compdateactual -ne $null ) -and (( get-date($compdateactual)  -Format "MM-dd-yyyy" ) -ne '01-01-0001')) {
          $compdate =   get-date($compdateactual)  -Format "MM-dd-yyyy" 
          $comptime = get-date($compdateactual)  -Format "hh:mm:ss"
      }
      else{
           $compdate=  $null
            $comptime = $null
      }
      if ( $esgitem.'Total Survey Time' -gt 0)
      {
          $totaltime =  $esgitem.'Total Survey Time' + ' sec'
      }
      else
      {
          $totaltime = " "
      }
    $surveydata = [ordered]@{
      'Language' = $esgitem.Language.Name
      'Survey Type' = $esgitem.'Survey Type'
      'Survey Name' =  $esgitem.'Survey Name'
      'Survey Form ID' =  $esgitem.'Survey Form ID'
      'User Email' =  $esgitem.'User Email'
      'Session ID' = $esgitem.'Session ID'
      'IsCompleted' =  $esgitem.'IsCompleted'
      'Current Panel Count' =  $esgitem.'Current Panel Count'
      'Start Date' =  $stdate
      'Start Time' = $sttime
      'Completion Date' =  $compdate
      'Completion Time' = $comptime
      'Total Survey Time' = $totaltime
        }
   $myObject = [pscustomobject]$surveydata
   $allpanels = $inputlist | select -property panelname -unique
   $panelnumber=0
   $panelresult = $null
   $allpanels | foreach{
       $panelname= $_.PanelName   
       $panelnumber = $panelnumber + 1
       
       #to fetch specific panel details
       $paneltimespent=$timelist  | where {$_.panelname -eq $panelname  }
       $panelresult = $resultlist | where {$_.panelname -eq $panelname  }
       $panelinput = $inputlist | where {$_.panelname -eq $panelname  }
       
       try {
       #to add panel name
       $myObject | Add-Member -Name ('P' + $panelnumber + '_PanelName') -Type NoteProperty -Value $panelname
       #to add timespent
       $paneltimespent | foreach{
            $timespent = $_
            $myObject | Add-Member -Name ('P' + $panelnumber + '_TimeSpent') -Type NoteProperty -Value $timespent.TimeSpent
       }  # foreach  add timespent
      
        #to add result data
       $panelresult | foreach{
            $result = $_
            $myObject | Add-Member -Name ('P' + $panelnumber + '_AverageScore') -Type NoteProperty -Value $result.AverageScore
             $myObject | Add-Member -Name ('P' + $panelnumber + '_EvaluatedScore') -Type NoteProperty -Value $result.EvaluatedScore
              $myObject | Add-Member -Name ('P' + $panelnumber + '_Rating') -Type NoteProperty -Value $result.Rating
       } #foreach add result data
      
       # to add que/ans detail
        $questionNo=0
        $panelinput | foreach{
            $inputdetail = $_
            $questionNo = $questionNo + 1
            $myObject | Add-Member -Name ('P' + $panelnumber + '_Q' + $questionNo + '_QuestionKey') -Type NoteProperty -Value $inputdetail.QuestionKey
             $myObject | Add-Member -Name ('P' + $panelnumber + '_Q' + $questionNo + '_QuestionText') -Type NoteProperty -Value $inputdetail.QuestionText
              $myObject | Add-Member -Name ('P' + $panelnumber + '_Q' + $questionNo + '_AnswerText') -Type NoteProperty -Value $inputdetail.AnswerText
               $myObject | Add-Member -Name ('P' + $panelnumber + '_Q' + $questionNo + '_CodeValue') -Type NoteProperty -Value $inputdetail.CodeValue
       }  # foreach  add que/ans detail
       
      
       
       } #try end
       catch {
           write-warning "Error in writing JSON data." 
           write-warning $Error[0]
            continue;
       }
     } #of panel
   $customList.Add($myObject)
    }#childitem of esg
     $customList | show-listview
} #$surveyFormFolders

}
catch {
    write-warning $Error[0]
    continue;
}

$StopWatch.Stop();

write-output "Total Execution Time: $($StopWatch.Elapsed.ToString())";