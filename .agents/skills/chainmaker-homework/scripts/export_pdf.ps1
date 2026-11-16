param([Parameter(Mandatory=$true)][string]$Docx)
$ErrorActionPreference='Stop'
$pdf=[IO.Path]::ChangeExtension($Docx,'.pdf')
if(Test-Path -LiteralPath $pdf){throw 'PDF already exists'}
$word=$null; $document=$null
try {
 $word=New-Object -ComObject Word.Application
 $word.Visible=$false
 $document=$word.Documents.Open($Docx,$false,$true)
 $document.ExportAsFixedFormat($pdf,17)
} finally {
 if($document){$document.Close(0)}
 if($word){$word.Quit()}
}
