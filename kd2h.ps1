# Convert .doc to .html  
param([string]$docpath,[string]$htmlpath = $docpath)  
$srcfiles = Get-ChildItem -Path $docPath -filter "*.doc"  
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatFilteredHTML");  
$word = new-object -comobject word.application  
$word.Visible = $False  

function saveas-filteredhtml{  
  $name = $doc.basename  
  $savepath = "$htmlpath\" + $name + ".html"  
  write-host $name  
  Write-Host $savepath  
  $opendoc = $word.documents.open($doc.FullName);  
  $opendoc.saveas([ref]$savepath, [ref]$saveFormat);  
  $opendoc.close();  
}  

ForEach ($doc in $srcfiles) {  
  Write-Host "Processing :" $doc.FullName  
  saveas-filteredhtml  
}  
$word.quit();  
