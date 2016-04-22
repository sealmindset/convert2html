$files = gci -filter "*.docx"
$savedir = "F:\ConvertedFiles\"
$word = New-Object -ComObject "Word.Application"
$word.Visible = $true

foreach ($document in $files) {
  write-host $document.Name
  $existingDoc=$word.Documents.Open($document.FullName)
  $name = $document.Name.Replace(" ","_");
  $saveaspath = $savedir + $name.Replace('.docx','.htm')
  $wdFormatHTML = [ref] 10
  $existingDoc.WebOptions.AllowPNG = $true
  $existingDoc.SaveAs( [ref] $saveaspath,$wdFormatHTML )
  $existingDoc.Close()
}

$Word.Quit() 
