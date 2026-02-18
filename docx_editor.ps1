# para atualizar vários documentos
# altera palavras da lista em todos os docx da pasta baseado no $myList



$myList = @{
	"foo" = "bar";
	"foo@hello.world" = "bar@hello.world";
}

$filePrefix = "_autoupdated_"

function Init-Word() {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $false
	return $word
}

function Open-File($word, $filePath) {
	$doc = $word.Documents.Open($filePath)
	return $doc
}

function wordReplace($doc, $FindText, $ReplaceText){
	$MatchCase = $false
	$MatchWholeWorld = $true
	$MatchWildcards = $false
	$MatchSoundsLike = $false
	$MatchAllWordForms = $false
	$Forward = $false
	$Wrap = 1
	$Format = $false
	$Replace = 2

    $doc.Content.Find.Execute($FindText, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $ReplaceText, $Replace) | Out-Null
}

function Replace-Words($doc, $hashTable) {
	foreach ($key in $hashTable.Keys) {
		$cur = $key
		$new = $hashTable[$key]
		wordReplace $doc $cur $new
	}
}

function Save-And-Quit($word, $doc, $newName) {
	$doc.SaveAs([ref]"$PSScriptRoot\$filePrefix$newName.docx")
	$doc.Close(-1)
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
}



function Main {
	Write-Host "Iniciando" -Foreground Green
	Write-Host "" -Foreground Green
	
	$word = Init-Word
	Get-ChildItem "$PSScriptRoot" -Filter *.docx | 
	Foreach-Object {
		$nome = $_.Name
		$nomeSemDocx = $_.BaseName
		$path = $_.FullName
		
		if ($nome -like "$filePrefix*") {
			Write-Host "$nome ignorado" -Foreground Yellow
			Write-Host ""
			Return 
		}
		
		Write-Host "Editando " -NoNewLine
		Write-Host "$nome" -Foreground Yellow
		
		Write-Host "Abrindo arquivo"
		$doc = Open-File $word $path
		
		Write-Host "Substituindo palavras"
		Replace-Words $doc $myList
		
		Write-Host "Salvando... " -NoNewLine
		Save-And-Quit $word $doc $nomeSemDocx
		Write-Host "OK"
		
		Write-Host ""
	}
	
	$word.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
	
	Write-Host "Finalizado" -Foreground Green
}



Main
