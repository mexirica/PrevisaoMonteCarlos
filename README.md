# PrevisaoMonteCarlos
Programa feito durante faculdade. Programa preve preço das ações usando o metodo estatistico de monte carlo.

![Image](https://user-images.githubusercontent.com/67772460/198433077-66f70913-70a2-43d3-a8b4-e314cfd3caf5.png)


Public parar As Integer
Public gf As Integer

Sub CapturaDados()
Dim Waitsec As Single
Dim WSD As Worksheet
Dim WSW As Worksheet
Dim connectstring As String
Dim linhafinal As Long
Dim proxlinha As Long
Dim linharesfinal As Long
Dim i As Integer
Dim j As Integer
Dim alpha As Single
Dim betha As Single
Dim n As Integer
Dim medx As Single
Dim medy As Single
Dim x(100) As Integer
Dim y(100) As Single
Dim somax As Single
Dim somax2 As Double
Dim somay As Single
Dim somaxy As Single
Dim prev As Single
Dim desvp As Single
Dim somad As Single
Dim max As Single
Dim min As Single


Set WSD = Worksheets("portfolio")
Set WSW = Worksheets("workspace")
Waitsec = UserForm1.TextBox3.Value
NameProc = "Capturadados"
linhafinal = WSD.Cells(65536, 1).End(xlUp).Row
proxlinha = linhafinal + 1
Cells(2, 1) = UserForm1.TextBox1.Text

connectstring = "URL;https://www.bussoladoinvestidor.com.br/cotacao/" & WSD.Cells(2, 1).Text & ".asp"

For Each QT In WSW.QueryTables
   QT.Delete
Next QT

Set QT = WSW.QueryTables.Add(Connection:=connectstring, Destination:=WSW.Range("A1"))

With QT
        .Name = WSD.Cells(2, 1).Text
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlSpecifiedTables
        .WebFormatting = xlWebFormattingNone
        .WebTables = "8"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
        
    WSW.Cells(1, 3) = Right(WSW.Cells(1, 2), 5)
    WSD.Cells(proxlinha, 1) = WSW.Cells(1, 3).Value * 1
 
     linharesfinal = WSW.Cells(65536, 1).End(xlUp).Row
     
 For i = 1 To linharesfinal
   For j = 1 To 15
      WSW.Cells(i, j).EntireRow.Delete
        Next j
  Next i
  

'+++++++++++++++++++++++++++++++ programa funcionando ++++++++++++++
If parar = 1 Then
NextTime = Time + TimeSerial(0, 0, Waitsec)
'++++++++++++++++++++++++++roda o relogio a cada próxima atualização

Application.OnTime earliesttime:=NextTime, procedure:=NameProc
Application.Wait (Now + TimeValue("0:00:05"))

'++++++++++++++++++++++++++++++ parada do programa +++++++++++++++++

ElseIf parar = 0 Then
On Error Resume Next
End If

If linhafinal >= 3 Then
n = Cells(100000, 1).End(xlUp).Row - 2
For i = 1 To n
    x(i) = i
    y(i) = Cells(i + 2, 1)
Next i
'''''''''''''''''''''''''''''''''''SOMAS'''''''''''''''''''''''''''''''''''''''''''''''''
somax2 = 0
somax = 0
somay = 0
somaxy = 0
For i = 1 To n
    somax = somax + x(i)
    somay = somay + y(i)
    somax2 = somax2 + (x(i) ^ 2)
    somaxy = somaxy + (x(i) * y(i))
Next i
''''''''''''''''''''''''''''Médias''''''''''''''''''''''''''''''''''''''''''''''''''''''
medx = somax / n
medy = somay / n
''''''''''''''''''''Cálculo de alpha e betha''''''''''''''''''''''''''''''''''''''''''''

betha = ((n * somaxy) - (somax * somay)) / ((n * somax2) - (somax ^ 2))
alpha = medy - (betha * medx)

'''''''''''''''''''''' a= e b= '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
UserForm1.TextBox4.Value = alpha
UserForm1.TextBox5.Value = betha
'''''''''''''''''''''''''''previsão''''''''''''''''''''''''''''''''''''''''''''''''''''''
prev = alpha + betha * x(n)
UserForm1.TextBox6.Value = prev

'''''''''''''''''''''''''opções de compra e venda''''''''''''''''''''''''''''''''''''''''
If y(n) > prev Then
    UserForm1.TextBox7 = "Acima"
    UserForm1.TextBox8 = "Vender"
    Else
    UserForm1.TextBox7 = "Abaixo"
    UserForm1.TextBox8 = "Comprar"
End If
''''''''''''''''''''''''''''''''''''''média do preço'''''''''''''''''''''''''''''''''''''
UserForm1.TextBox9.Value = medy
''''''''''''''''''''''''''''''Calcular o desvio padrão'''''''''''''''''''''''''''''''''''
somad = 0
For i = 1 To n
somad = somad + (y(i) - medy) ^ 2
Next i

desvp = Sqr((somad / n))
UserForm1.TextBox10.Value = desvp
UserForm1.TextBox11.Value = betha
''''''''''''''''''''''''''''Valor Máximo e Mínimo''''''''''''''''''''''''''''''''''''''''
max = y(1)
For i = 1 To n
    If y(i) > max Then
    max = y(i)
    End If
Next i

min = y(1)
For i = 1 To n
    If y(i) < min Then
    min = y(i)
    End If
Next i

UserForm1.TextBox12.Value = max
UserForm1.TextBox13.Value = min
UserForm1.TextBox2.Value = y(n)

''''''''''''''''''''''''''''Atualização do gráfico''''''''''''''''''''''''''''''''''''''''
If gf = 2 Then
    grafico2
End If

End If


End Sub
Sub criargrafico()
Dim n As Single

    n = Cells(65536, 1).End(xlUp).Row
    Range(Cells(3, 1), Cells(n, 1)).Select
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    ActiveChart.SetSourceData Source:=Range(Cells(3, 1), Cells(n, 1))
    ActiveSheet.ChartObjects(1).Activate

End Sub
Sub grafico2()
Dim n As Single
    
    n = Cells(65536, 1).End(xlUp).Row
    Range(Cells(3, 1), Cells(n, 1)).Select
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.SetSourceData Source:=Range(Cells(3, 1), Cells(n, 1))
    
    
End Sub
Sub apagargrafico()

    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Parent.Delete
    
End Sub
Sub botao()

UserForm1.Show


End Sub


-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub CommandButton1_Click()
    '
    'botão iniciar

parar = 1
Call CapturaDados


End Sub

Private Sub CommandButton2_Click()
'
'
'botão parar

parar = 0

End Sub

Private Sub CommandButton3_Click()
'
    'Criar gráfico
    gf = 1
    Call criargrafico
    gf = 2


End Sub

Private Sub CommandButton4_Click()

'
' apagar grafico
'
  
Call apagargrafico


End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub
