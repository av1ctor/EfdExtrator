$documents_path = Split-Path -parent $MyInvocation.MyCommand.Path

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object -ThrottleLimit 10 -Parallel {
	$word_app = New-Object -ComObject Word.Application
	$word_app.Visible = $False

    $document = $word_app.Documents.Open($_.FullName)
    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
	
	$word_app.Quit()
}
