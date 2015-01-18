<#
To Do: use jsonObject = [ordered]@{} to automate JSON syntax - either all Convertto-JSON or JavaScript.Serializer

#>


$ErrorActionPreference = "Stop"
try {   
   	$libpath = Split-Path $MyInvocation.MyCommand.Path -Parent
   	write-host "`nLoading iTextSharp.dll $libpath for PDF reading `n"
   	Add-Type -Path "$libpath\lib\iTextSharp.dll"
   	[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
   	write-host Loading System.Web for URL Encoding
     }

Catch { write-host "Error - Stopping Script: `n$_" ; exit}

$bc = [char]123	#{ 
$ec = [char]125	#} 
$bb = [char]91	#[  
$eb = [char]93	#]  
$n = [char]58	#:  
$q = [char]34	#"  
$c = [char]44	#,  
$lb = "`n"	# LinFeed
$tb = "`t"	# Tab
$errors = 0
$filesparsed = 0
$totalpages = 0
$files = Get-Childitem . -filter "*.pdf"
$totalfiles = $files.Count
$filesparsed = $totalfiles
$text = ""
foreach ($file in $files) {
	try
		{
		   Write-Host ""
		   write-host "	Creating Reader Object on file $file"
		   $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file.fullname
		   
		   $pdfinfo = 	$bc + $lb +
		                " " + $q + "PDFDocument" + $q + $n + $q + $file.name + $q + $c + $lb +
		                " " + $q + "DocuInfo" + $q + $n + $bb + $lb +
		                "  " + $bc + $lb +
		                "   " + $q + "FileName" + $q + $n + $q + $file.name + $q + $c + $lb +
		                "   " + $q + "FileLength" + $q + $n + $reader.FileLength + $c + $lb +
		                "   " + $q + "Tampered" +  $q + $n + $q + $reader.Tampered + $q + $c + $lb +
		                "   " + $q + "PdfVersion" +  $q + $n + $q + $reader.PdfVersion + $q + $c + $lb +
		                "   " + $q + "Appendable" +  $q + $n + $q + $reader.Appendable + $q + $c + $lb +
		                "   " + $q + "Author" +  $q + $n + $q + $reader.info.Author + $q + $c + $lb +
		                "   " + $q + "CreationDate" +  $q + $n + $q + $reader.info.CreationDate + $q + $c + $lb +
		                "   " + $q + "Creator" +  $q + $n + $q + $reader.info.Creator + $q + $c + $lb +
		                "   " + $q + "ModDate" +  $q + $n + $q + $reader.info.ModDate + $q + $c + $lb +
		                "   " + $q + "Producer" +  $q + $n + $q + $reader.info.Producer + $q + $c + $lb +
		                "   " + $q + "Title" +  $q + $n + $q + $reader.info.Title + $q + $c + $lb +
		                "   " + $q + "TotalPages" +  $q + $n + $reader.NumberOfPages + $c + $lb +
		                "   " + $q + "Body" +  $q + $n + $bb + $lb 
		   $pdfinfo | Out-File "$file.json"

		   for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
		    		Write-Host "		Reading Page $page"
				$strategy = new-object  'iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy'
		  		$currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy);
		  		[string[]]$stringText = [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default, [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));
		  		[string[]]$Text += [system.text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert( [system.text.encoding]::default, [system.text.encoding]::UTF8, [system.text.Encoding]::Default.GetBytes($currentText)));
				$htmlText = [System.Web.HttpUtility]::UrlEncode($stringText) 
		  		$pdfPageText =   "    " + $bc + $lb + 	 
		  		                 "     " + $q + "page" + $q + $n + $page + $c + $lb + 
		                                 "     " + $q + "pagetext" + $q + $n + $q + $htmlText + $q + $lb + 
		                                 "    " + $ec + $lb  		                                
		  		if($page -ne $Reader.NumberOfPages){$pdfPageText +=$c}
		         	$pdfPageText | Out-File "$file.json" -append -noclobber
		         	$totalpages++
		  	}
		  	
		   $PDFInfo | Out-File "$file.txt" 
		   $text | Out-File "$file.txt" -append -noclobber
		   $text=""		 
		}
	catch
		{
			Write-Host "	Error: Unable to create reader -  $_"
			$errors++
			$filesparsed = $filesparsed -1
		}
	finally
		{
			$closejson =  "   " + $eb + $lb + 
		                      "  " + $ec + $lb + 
		                      " " + $eb + $lb + 
		                      $ec
		        $closejson | Out-File "$file.json" -append -noclobber
			Write-Host "	Completed Reading $file"
			Write-host ""
			Write-Host $("#" * 150)
		}
	}
Write-Host ""
Write-Host $("-" * 25)
Write-Host "Totals Files: $totalfiles"	
Write-Host "Files Parsed: $filesparsed"
Write-Host "Pages Parsed: $totalPages"
Write-Host "Errors: $errors"
Write-Host $("-" * 25)
Write-Host ""	
$Reader.Close();



<#
json schema
{
  "PDFDocument": "name",
  "DocInfo": [
    {
      "Filename": "name",
      "Author": "name",
      "CreationDate": "date",
      "Creator": "",
      "ModificationDate": "",
      "Producer": "",
      "Title": "",
      "TotalPages": 1,
      "body": [
        {
          "page": 1,
          "pagetext": "text"
        },
        {
          "page": 2,
          "pagetext": "text"
        },
        {
          "page": 3,
          "pagetext": "text"
        }
      ]
    }
  ]
}
#>