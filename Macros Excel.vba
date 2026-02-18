Sub ConverterCSV()
'
' Converte csv de uma linha em tabela
' Colunas onde os 9 primeiros valores são data são automaticamente considerados como data
'
    
    Dim primeiraLinha As String
    Dim qtdeColunas As Long
    Dim fieldInfo() As Variant
    Dim ultimaColuna As Long
    Dim linha As Long
    Dim coluna As Long
    Dim éData As Boolean
    Dim nomeColuna As String
    
 
    ' para deixar macro mais rápido
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    
    primeiraLinha = ActiveSheet.Range("A1").Text
    qtdeColunas = UBound(Split(primeiraLinha, ",")) + 1
    ReDim fieldInfo(1 To qtdeColunas)

    For i = 1 To qtdeColunas
        fieldInfo(i) = Array(i, xlTextFormat)
    Next i

    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=True, Space:=False, Other:=False, TrailingMinusNumbers:=True, _
        fieldInfo:=fieldInfo
    
    ultimaColuna = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    For coluna = 1 To ultimaColuna
        éData = True
        For linha = 2 To 10
            
            If Not IsDate(ActiveSheet.Cells(linha, coluna).Value) Then
                éData = False
            End If
        Next linha
        
        If éData Then
            nomeColuna = Split(Cells(1, coluna).Address(True, False), "$")(0)
            ConverterTextoParaData nomeColuna
        End If
    Next coluna

    Range("A1").Select
    Application.Wait (Now + TimeValue("0:00:01"))
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ' Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "Tabela1"
    Range("Tabela1[#All]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Columns.AutoFit
    Range("A1").Select
    ActiveWindow.DisplayGridlines = False
    
    Call SubstituirCaracteresErrados
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

Sub ConverterTextoParaData(coluna As String)
    Dim celula As Range
    Dim linhaMax As Long
    
    linhaMax = ActiveSheet.Cells(ActiveSheet.Rows.Count, coluna).End(xlUp).Row
    
    
    For i = 1 To linhaMax
        Set celula = ActiveSheet.Cells(i, coluna)
        If Not IsError(celula.Value) Then
            If IsDate(celula.Value) Then
                celula.Value = CDbl(CDate(celula.Value))
                celula.NumberFormat = "dd/mm/yyyy HH:mm"
            End If
        End If
    Next i
End Sub

Sub SubstituirCaracteresErrados()
'
' Troca caracteres especiais incorretos pelos corretos.
'

    ActiveSheet.Cells.Replace What:="Ã©", Replacement:="é", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveSheet.Cells.Replace What:="Ã§Ã£", Replacement:="çã", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveSheet.Cells.Replace What:="â€“", Replacement:="-", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveSheet.Cells.Replace What:="Ã§", Replacement:="ç", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveSheet.Cells.Replace What:="Ãµ", Replacement:="õ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveSheet.Cells.Replace What:="Ã”", Replacement:="Ô", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ActiveSheet.Cells.Replace What:="Ã­", Replacement:="í", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    ' ActiveSheet.Cells.Replace What:="", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
 
End Sub

Sub BparaGB()
'
' Converte de bytes para gigabytes
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim celula As Range
    Dim ultimaLinha As Long
    Dim colunaSelecionada As Long
    
    ' Defina a planilha ativa
    Set ws = ActiveSheet
    
    ' Verifique se há uma seleção
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione uma célula em uma coluna com valores em bytes para converter.", vbExclamation
        Exit Sub
    End If
    
    ' Defina a coluna selecionada
    colunaSelecionada = Selection.Column
    
    ' Encontre a última linha com dados na coluna selecionada
    ultimaLinha = ws.Cells(ws.Rows.Count, colunaSelecionada).End(xlUp).Row
    
    ' Defina o intervalo de células para a coluna selecionada
    Set rng = ws.Range(ws.Cells(1, colunaSelecionada), ws.Cells(ultimaLinha, colunaSelecionada))
    
    ' Iterar sobre cada célula na coluna selecionada
    For Each celula In rng
        If IsNumeric(celula.Value) And celula.Value > 1073741824 Then
            ' Converte bytes para gigabytes (1 GB = 2^30 bytes) e arredonda para 2 casas decimais
            celula.Value = Round(celula.Value / (1024 ^ 3), 2)
        Else
            ' Deixa como está
        End If
    Next celula
    
End Sub

Sub OrdenarIPs()
'
'   Ordena coluna selecionada por IP
'   Só funciona se estiver em uma tabela
'
    Dim cel As Range
    Dim ColunaOriginal As Range
    Dim ColunaNova As Range
    
    rowNumber = ActiveSheet.UsedRange.Rows.Count
    Set ColunaOriginal = Range(Cells(2, Selection.Column), Cells(rowNumber, Selection.Column))
    Selection.Offset(0, 1).EntireColumn.Insert
    Cells(1, Selection.Column + 1).Value = "Sorted IP"
    Set ColunaNova = Range(Cells(2, Selection.Column + 1), Cells(rowNumber, Selection.Column + 1))
    
    For Each cel In ColunaNova.Cells
        cel.Value = IPToInteger(cel.Offset(0, -1).Value)
    Next cel
    
    ColunaNova.NumberFormat = "0"
    ColunaNova.Sort Key1:=ColunaNova, Order1:=xlAscending, Header:=xlYes
    
    Columns(Selection.Column + 1).Delete
End Sub

Function IPToInteger(ip As String) As String
    Dim octets() As String
    Dim partialResult As Long
    Dim result As String
    Dim i As Long
    
    If IsNumeric(ip) And InStr(ip, ".") = 3 Then
        octets = Split(ip, ".")
        partialResult = 0
        For i = 0 To 3
            partialResult = partialResult + CLng(octets(i)) * (256 ^ (3 - i))
        Next i
        result = CStr(partialResult)
    Else
        result = ip
    End If

    IPToInteger = result
End Function

Sub PontuarIPs()
'
'   Corrige IPs que Excel altera para número inteiro, adicionando pontos
'   Formato: XX.YYY.ZZZ.WWW
'
    Dim cell As Range
    Dim ip As String
    Dim ipFormatado As String
    
    ' para deixar macro mais rapida
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            ip = CStr(cell.Value)
            If Len(ip) = 11 And CheckValidOctet(ip) Then
                If IsNumeric(Left(ip, 2)) And IsNumeric(Mid(ip, 3, 3)) And IsNumeric(Mid(ip, 6, 3)) And IsNumeric(Mid(ip, 9, 3)) Then
                    ipFormatado = Left(ip, 2) & "." & Mid(ip, 3, 3) & "." & Mid(ip, 6, 3) & "." & Mid(ip, 9, 3)
                    cell.Value = ipFormatado
                End If
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True


End Sub

Function CheckValidOctet(ip As String) As Boolean
    firstOctet = Left(ip, 2)
    secondOctet = Mid(ip, 3, 3)
    thirdOctet = Mid(ip, 6, 3)
    fourthOctet = Right(ip, 3)
    
    If IsNumeric(firstOctet) And IsNumeric(secondOctet) And IsNumeric(thirdOctet) And IsNumeric(fourthOctet) Then
        If firstOctet = "10" And secondOctet = "162" Then
            If CInt(thirdOctet) > 0 And CInt(thirdOctet) <= 255 And CInt(fourthOctet) > 0 And CInt(fourthOctet) <= 255 Then
                CheckValidOctet = True
            Else
                CheckValidOctet = False
            End If
        Else
            CheckValidOctet = False
        End If
    Else
        CheckValidOctet = False
    End If
End Function

Sub RemoverLinhasDuplicadas()
'
'   Remove linhas que tem valores duplicados
'
    Dim rngColuna As Range
    Dim cel As Range
    Dim ultimaLinha As Long
    Dim i As Long
    Dim corDuplicado As Long
    Dim totalAntes As Long
    Dim removidas As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Define a cor usada na formatação condicional para duplicados
    corDuplicado = RGB(255, 0, 0)
    
    ' Garante que apenas uma coluna esteja selecionada
    If Selection.Columns.Count > 1 Then
        MsgBox "Por favor, selecione apenas uma coluna.", vbExclamation
        Exit Sub
    End If
    
    ' Define o intervalo da coluna até a última célula com dados
    ultimaLinha = Cells(Rows.Count, Selection.Column).End(xlUp).Row
    totalAntes = ultimaLinha
    Set rngColuna = Range(Cells(1, Selection.Column), Cells(ultimaLinha, Selection.Column))
    
    ' Aplica formatação condicional para duplicados
    With rngColuna.FormatConditions
        .Delete
        .AddUniqueValues
        With .Item(1)
            .DupeUnique = xlDuplicate
            .Font.Color = corDuplicado
        End With
    End With

    ' Aguarda atualização visual
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Varre de baixo para cima, removendo linhas com célula duplicada (vermelha)
    removidas = 0
    For i = ultimaLinha To 1 Step -1
        Set cel = Cells(i, Selection.Column)
        If cel.DisplayFormat.Font.Color = corDuplicado Then
            Rows(i).Delete
            removidas = removidas + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' Exibe resultado
    MsgBox "Total de linhas antes: " & totalAntes & vbCrLf & _
           "Linhas removidas: " & removidas & vbCrLf & _
           "Total após remoção: " & (totalAntes - removidas), vbInformation, "Resumo da Limpeza"
End Sub




