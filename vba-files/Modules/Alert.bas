Attribute VB_Name = "OpenFiles"

sub AbrirArquivos()

Dim objShell As Object

Set objShell = CreateObject("Shell.Application")

caminho_pasta = "F:\cursos\HASHTAG\VBA\Abrindo qualquer arquivo ou pasta"

objShell.Open(caminho_pasta)

msgBox "Pasta aberta"


end sub



