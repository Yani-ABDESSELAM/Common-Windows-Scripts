#### Author Allenage.com ###

### Get Windows Update status ###


#### Enter Computer Name On Prompt to get installed Updates #####

$computers=[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null  
$computers= [Microsoft.VisualBasic.Interaction]::InputBox("Enter Computername")
foreach ($computer in $computers){
if(!(Test-Connection -Cn $computer -BufferSize 16 -Count 1 -ea 0 -quiet))
{write-host "cannot reach $computer" -f red}

else {$Session = New-Object -ComObject "Microsoft.Update.Session"

$Searcher = $Session.CreateUpdateSearcher()

$historyCount = $Searcher.GetTotalHistoryCount()

$Searcher.QueryHistory(0, $historyCount) | Select-Object Date,

   @{name="Operation"; expression={switch($_.operation){

       1 {"Installation"}; 2 {"Uninstallation"}; 3 {"Other"}}}},

   @{name="Status"; expression={switch($_.resultcode){

       1 {"In Progress"}; 2 {"Succeeded"}; 3 {"Succeeded With Errors"};

       4 {"Failed"}; 5 {"Aborted"}

}}}, Title | Export-Csv -NoType "$Env:userprofile\Desktop\$computer-WindowsUpdates.csv"}}



### get Local computers Windows Updates ######

$Session = New-Object -ComObject "Microsoft.Update.Session" 
 
$Searcher = $Session.CreateUpdateSearcher() 
 
$historyCount = $Searcher.GetTotalHistoryCount() 
 
$Searcher.QueryHistory(0, $historyCount) | Select-Object Date, 
 
   @{name="Operation"; expression={switch($_.operation){ 
 
       1 {"Installation"}; 2 {"Uninstallation"}; 3 {"Other"}}}}, 
 
   @{name="Status"; expression={switch($_.resultcode){ 
 
       1 {"In Progress"}; 2 {"Succeeded"}; 3 {"Succeeded With Errors"}; 
 
       4 {"Failed"}; 5 {"Aborted"} 
 
}}}, Title | Export-Csv -NoType "$Env:userprofile\Desktop\WindowsUpdates.csv"

##### Remote Computers with server or ComputerName on Txt file ##

$computers = get-Content c:\computers.txt
foreach ($computer in $computers){
if(!(Test-Connection -Cn $computer -BufferSize 16 -Count 1 -ea 0 -quiet))
{write-host "cannot reach $computer" -f red}

else {$Session = New-Object -ComObject "Microsoft.Update.Session"

$Searcher = $Session.CreateUpdateSearcher()

$historyCount = $Searcher.GetTotalHistoryCount()

$Searcher.QueryHistory(0, $historyCount) | Select-Object Date,

   @{name="Operation"; expression={switch($_.operation){

       1 {"Installation"}; 2 {"Uninstallation"}; 3 {"Other"}}}},

   @{name="Status"; expression={switch($_.resultcode){

       1 {"In Progress"}; 2 {"Succeeded"}; 3 {"Succeeded With Errors"};

       4 {"Failed"}; 5 {"Aborted"}

}}}, Title | Export-Csv -NoType "$Env:userprofile\Desktop\$computer-WindowsUpdates.csv"}}