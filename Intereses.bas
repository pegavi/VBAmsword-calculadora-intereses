Attribute VB_Name = "Intereses"
Option Explicit

Private Type PeriodosDeIntereses
    fechaInicio As Date
    fechaFinal As Date
    Tipo As Double
    Dias As Integer
    InteresesPeriodo As Double
    
End Type

'Sub test()
'
'Dim myFile As String, text As String, textline As String, posLat As Integer, posLong As Integer
'
'Dim rutaDirectorio As String
'
'rutaDirectorio = Left(ThisDocument.FullName, InStrRev(ThisDocument.FullName, "\"))
'
'myFile = rutaDirectorio & "hola.txt"
'
'Open myFile For Input As #1
'
'Do Until EOF(1)
'    Line Input #1, textline
'    text = text & textline
'Loop
'
'Close #1
'
'MsgBox text
'
'End Sub


'Sub test()
'
'Dim parse() As String
'
'parse = Split("1/1/2010:4:1/1/2011:4:1/1/2012:4:1/1/2013:4:1/1/2014:4:1/1/2015:3,5:1/1/2016:3:1/1/2017:3:1/1/2018:3:1/1/2019:3:1/1/2020:3:1/1/2021:3:1/1/2022:3:1/1/2023:3,25:31/12/2023", ":")
'
'Call insertar(DateSerial(2024, 2, 2), DateSerial(2024, 7, 9), CDbl(1000), parse, False)
'
'
'
'End Sub

Sub CalcularIntereses()



If Selection.StoryType <> wdMainTextStory Then

    MsgBox "La selección debe estar en la sección principal del documento (no notas al pie, encabezados...)"
    
    Exit Sub

ElseIf Selection.Information(wdWithInTable) = True Then

    MsgBox "La selección no puede ser una tabla"
    Exit Sub
    
ElseIf Selection.Type <> wdSelectionIP And Selection.Type <> wdSelectionNormal Then

    MsgBox "Selección no válida. Selecciona el lugar donde quieres que se introduzca la tabla con el cálculo de intereses."
    Exit Sub

Else


    Selection.MoveLeft (1)

    If Selection.Information(wdWithInTable) Then
       
    
    Selection.MoveRight (1)
    Selection.InsertAfter (vbCr) 'ANSI 13
    Selection.Collapse (wdCollapseEnd)
    
    
    End If
    
    Selection.MoveRight (1)

    Call frmIntereses.Show
    'Call insertar(#5/15/2017#, #5/15/2025#, 123)

End If

End Sub

Sub insertar(fechaInicio As Date, fechaFinal As Date, capital As Double, parse() As String, periodos As Boolean)

Dim a() As PeriodosDeIntereses
Dim i As Integer
Dim e As Integer
Dim rng As Range
Dim tbl As Table
Dim total As Double
Dim ajusteNecesario As Boolean

a = Intereses(fechaInicio, fechaFinal, capital, parse)

For i = 0 To UBound(a)
    
        total = total + a(i).InteresesPeriodo
    
    Next i
    

Set rng = Selection.Range

If periodos Then
       
       
    Set tbl = rng.Tables.Add(rng, UBound(a) + 3, 6)
    tbl.Borders.Enable = True
    tbl.AllowAutoFit = True
    tbl.Range.Paragraphs.Alignment = wdAlignParagraphCenter
    tbl.Range.ParagraphFormat.SpaceAfter = 0
    tbl.Rows(1).Cells(1).Range.text = "Capital"
    tbl.Rows(1).Cells(2).Range.text = "Desde"
    tbl.Rows(1).Cells(3).Range.text = "Hasta"
    tbl.Rows(1).Cells(4).Range.text = "Días"
    tbl.Rows(1).Cells(5).Range.text = "Tipo"
    tbl.Rows(1).Cells(6).Range.text = "Total"
    tbl.Rows(1).Range.Font.Italic = True
    tbl.Rows(1).Range.Font.Bold = True
   
    
    For i = 2 To tbl.Rows.Count - 1
    
        tbl.Rows(i).Cells(1).Range.text = FormatCurrency(capital)
        
        If i = 2 Then
            tbl.Rows(i).Cells(2).Range.text = fechaInicio
        Else
            tbl.Rows(i).Cells(2).Range.text = a(i - 2).fechaInicio
        End If
        
        If i = tbl.Rows.Count - 1 Then
            tbl.Rows(i).Cells(3).Range.text = fechaFinal
        Else
            tbl.Rows(i).Cells(3).Range.text = a(i - 2).fechaFinal
        End If
        tbl.Rows(i).Cells(4).Range.text = a(i - 2).Dias
        tbl.Rows(i).Cells(5).Range.text = a(i - 2).Tipo & "%"
        tbl.Rows(i).Cells(6).Range.text = FormatCurrency(a(i - 2).InteresesPeriodo)
    
    Next i
 
    tbl.Rows(tbl.Rows.Count).Cells(5).Range.text = "TOTAL:"
    tbl.Rows(tbl.Rows.Count).Cells(6).Range.text = FormatCurrency(total)
    tbl.Rows(tbl.Rows.Count).Range.Font.Bold = True
Else

Set tbl = rng.Tables.Add(rng, 2, 5)
    tbl.AllowAutoFit = True
    tbl.Borders.Enable = True
    tbl.Range.ParagraphFormat.SpaceAfter = 0
    tbl.Rows(1).Cells(1).Range.text = "Capital"
    tbl.Rows(1).Cells(2).Range.text = "Desde"
    tbl.Rows(1).Cells(3).Range.text = "Hasta"
    tbl.Rows(1).Cells(4).Range.text = "Días"
    tbl.Rows(1).Cells(5).Range.text = "Total"
    tbl.Rows(1).Range.Font.Italic = True
    
       
    tbl.Rows(2).Cells(1).Range.text = FormatCurrency(capital)
    tbl.Rows(2).Cells(2).Range.text = fechaInicio
    tbl.Rows(2).Cells(3).Range.text = fechaFinal
    tbl.Rows(2).Cells(4).Range.text = DateDiff("d", fechaInicio, fechaFinal) + 1
    tbl.Rows(2).Cells(5).Range.text = FormatCurrency(total)
        
    

End If




'tbl.Range.ParagraphFormat.SpaceAfter = 0

End Sub


Function Intereses(inicioComputo As Date, finComputo As Date, capital As Double, parse() As String) As PeriodosDeIntereses()

' Variables (legal, morosidad en operaciones comerciales...) para guardar un string con los tipos y periodos de cada tipo de interés, así como fecha final sobre la que hay información
'Dim legal As String

' Variable para convertir el string en fechas de inicio de cada periodo y su tipo de interés
'Dim parse() As String
'Variable type privado donde se almacenará la información de fechas, tipos, días y intereses
Dim periodos() As PeriodosDeIntereses
'Variables para guardar la primera y la última fecha para la que hay información disponible
Dim limitePeriodo As Date
'Dim inicioPeriodo As Date
'Variables de control
Dim e As Integer
Dim i As Integer
Dim j As Integer

'Strings que contienen los datos de cada tipo de interés. Eventualmente, habría que leerlos de un archivo externo en lugar de tenerlos en una variable aquí.
'Incluye un par fecha:interés para cada periodo
'y una fecha final indicando el último día de aplicabilidad del último periodo sobre el que se tienen datos.
'legal = "1/1/2010:4:1/1/2011:4:1/1/2012:4:1/1/2013:4:1/1/2014:4:1/1/2015:3,5:1/1/2016:3:1/1/2017:3:1/1/2018:3:1/1/2019:3:1/1/2020:3:1/1/2021:3:1/1/2022:3:1/1/2023:3,25:31/12/2023"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Comprueba que la fecha de inicio no sea superior a la fecha de fin

If inicioComputo > finComputo Then

    MsgBox "La fecha de inicio (" & inicioComputo & ") no puede ser superior a la fecha de fin del cálculo (" & finComputo & ")."

    Exit Function

End If

'Carga la tabla de periodos y tipos en un array.
'Al comenzar en 0 el array, cada elemento par de "parse" es una fecha, y el impar siguiente es su tipo de interés.
'El último elemento Ubound es la fecha del último día de aplicación del último periodo registrado.
'parse = Split(legal, ":")

' La primera y la última fecha sobre la que hay datos válidos se asigna a una variable aparte.

limitePeriodo = parse(UBound(parse))
'inicioPeriodo = parse(0)

         
'OPCIONAL: para no cargar en periodos() todos los periodos, incluso los que no tienen intereses, calculamos qué índice de parse() contiene el primer periodo con intereses (i-2), y cuál el último (e-2)

For i = 0 To UBound(parse) - 1 Step 2

    If CDate(parse(i)) > inicioComputo Then
    
        Exit For
    
    End If

Next i

For e = 0 To UBound(parse) - 1 Step 2

    If CDate(parse(e)) > finComputo Then
    
        Exit For
    
    End If

Next e

If finComputo > limitePeriodo And inicioComputo < limitePeriodo Then

' Redimensionamo el array periodos()
ReDim periodos(((e + 2) - i) / 2)

Else

' Redimensionamo el array periodos()
ReDim periodos((e - i) / 2)

End If




' repasa el array "parse", tomando los elementos pares (i = 0, 2, 4...), que se corresponden con las fechas, y los asigna al array "periodos.FechaInicio".
' a continuación toma los elementos impares (i+1), correspondientes a los tipos de interés, y los asigna al array "periodos.Tipo".
' acaba en "Ubound(parse) - 1" para no incluir la fecha final (fin del último periodo sobre el que hay datos), que va en una variable aparte

If finComputo > limitePeriodo And inicioComputo > limitePeriodo Then

    With periodos(j)
        .fechaInicio = inicioComputo
        .Tipo = parse(UBound(parse) - 1)
        .fechaFinal = finComputo
        
    End With

Else

For i = i - 2 To e - 2 Step 2


    With periodos(j)
        .fechaInicio = parse(i)
        .Tipo = parse(i + 1)
        
        If (i + 2) < UBound(parse) Then

            .fechaFinal = CDate(parse(i + 2)) - 1
        Else

            .fechaFinal = parse(i + 2)

        End If
        
    End With
    
    
    j = j + 1

Next i

If j = UBound(periodos) Then

    With periodos(j)
        .fechaInicio = CDate(parse(UBound(parse))) + 1
        .Tipo = parse(UBound(parse) - 1)
        
        .fechaFinal = finComputo
        
    End With

End If

End If


For i = 0 To UBound(periodos)

    With periodos(i)

        If i = 0 Then

            If finComputo > .fechaFinal Then

                .Dias = DateDiff("d", inicioComputo, .fechaFinal) + 1
            Else

                .Dias = DateDiff("d", inicioComputo, finComputo) + 1

            End If

        Else

            .Dias = DateDiff("d", .fechaInicio, .fechaFinal) + 1

        End If
        
        .InteresesPeriodo = capital * .Tipo * .Dias / (DateDiff("d", DateSerial(year(.fechaInicio), 1, 1), DateSerial(year(.fechaInicio), 12, 31)) + 1) / 100

    End With


Next i

Intereses = periodos


End Function

