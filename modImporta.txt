Attribute VB_Name = "modImporta"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Macro desenvolvida por Marcelo Palladino em 07/02/2014 '
'                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim fileToOpen
Sub ImportarArquivo()

'Declara a variável para o nome do arquivo texto
Dim ArqBanco
ArqBanco = ""

'Abre janela para abertura do arquivo texto
fileToOpen = Application _
    .GetOpenFilename("Text Files (*.txt), *.txt") 'Define o tipo de arquivo permitido
If fileToOpen <> False Then
    'Se o arquivo for válido importa o novo arquivo
    ArqBanco = fileToOpen
End If
'Se nenhum arquivo for selecionado encerra
If fileToOpen = False Then End

'Define variáveis para abertura e leitura do arquivo texto
Dim sArquivo As String, sLinha As String, iARQ As Integer
'Nome do arquivo
sArquivo = ArqBanco
'Libera leitura do arquivo
iARQ = FreeFile
'Abre o arquivo
Open sArquivo For Input As iARQ

'Declara variáveis de linha e coluna para destino dos dados
Dim R, C
R = 1: C = 1
Cells(R, C).Activate

'Declara colunas dos dados e Delimitador de coluna
Dim sNome, sGenero, sNacional, sDelimitador
sDelimitador = "|" 'Nesse caso vamos usar o pipe"|"

'Declara as paradas. Como são três colunas, temos só duas paradas
Dim Parada1, Parada2
Parada1 = 0: Parada2 = 0

'Inicia o loop no arquivo texto
Do While Not EOF(iARQ)
    'Seleciona a célula
    Cells(R, C).Activate

    'Pega a linha atual do arquivo texto
    Line Input #iARQ, sLinha
    
    'Vamos definir as paradas "|"
    For i = 1 To Len(sLinha)
        If Parada1 = 0 And Mid(sLinha, i, 1) = "|" Then Parada1 = i
        If Parada1 > 0 And Mid(sLinha, i, 1) = "|" Then Parada2 = i
    Next i

    'Carregamos os dados para as constantes
    sNome = Mid(sLinha, 1, Parada1 - 1)
    sGenero = Mid(sLinha, Parada1 + 1, Parada2 - 1 - Parada1)
    sNacional = Mid(sLinha, Parada2 + 1, Len(sLinha))
    
    'Descarrega as constantes para a planilha
    Cells(R, C) = sNome
    Cells(R, C + 1) = sGenero
    Cells(R, C + 2) = sNacional
    
    'Avança uma linha
    R = R + 1
    
    'Limpa constantes
    sNome = ""
    sGenero = ""
    sNacional = ""
    Parada1 = 0
    Parada2 = 0
    
    DoEvents
Loop

End Sub
