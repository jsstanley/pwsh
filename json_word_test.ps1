# Basic Word doc creation
$Word = New-Object -ComObject Word.Application
$Word.Visible = $True
$Document = $Word.Documents.Add()
$Selection = $Word.Selection
$Selection.TypeText("Hello world")

# Using JSON data
$Response = Invoke-WebRequest -Uri 'https://jsonplaceholder.typicode.com/users' -UseBasicParsing
$Selection.TypeParagraph()
$Selection.TypeText("DATE: $(Get-Date)")
$Selection.TypeParagraph()
$Selection.TypeText($Response.Content)
$Document.PrintOut()
