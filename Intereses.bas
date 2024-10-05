Attribute VB_Name = "Intereses"
Option Explicit

Const MSG_SEL_MAIN_SECTION As String = "La selección debe estar en la sección principal del documento (no notas al pie, encabezados...)"
Const MSG_SEL_NO_TABLE As String = "La selección no puede ser una tabla"
Const MSG_SEL_INVALID As String = "Selección no válida. Selecciona el lugar donde quieres que se introduzca la tabla con el cálculo de intereses."

Private Type PeriodosDeIntereses
    fechaInicioPeriodo As Date
    fechaFinPeriodo As Date
    tipo As Double
    dias As Integer
    interesesPeriodo As Double
    
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

    ' Validación de selección
    If Not IsValidSelection() Then Exit Sub
    
    ' Ajustes para evitar que se fusione con una tabla anterior, si la selección está justo después de una tabla
    Call AdjustSelection
    
    ' Mostrar formulario de intereses
    frmIntereses.Show

End Sub

' Función para validar si la selección está en la Main Story, fuera de una tabla, y si hay algo seleccionado, para evitar errores.
Function IsValidSelection() As Boolean
    If Selection.StoryType <> wdMainTextStory Then
        MsgBox MSG_SEL_MAIN_SECTION
        IsValidSelection = False
    ElseIf Selection.Information(wdWithInTable) = True Then
        MsgBox MSG_SEL_NO_TABLE
        IsValidSelection = False
    ElseIf Selection.Type <> wdSelectionIP And Selection.Type <> wdSelectionNormal Then
        MsgBox MSG_SEL_INVALID
        IsValidSelection = False
    Else
        IsValidSelection = True
    End If
End Function

'Sirve para evitar que se fusione con una tabla anterior, si la selección está justo después de una tabla
Sub AdjustSelection()
    Selection.MoveLeft (1)
    If Selection.Information(wdWithInTable) Then
        Selection.MoveRight (1)
        Selection.InsertAfter (vbCr) 'ANSI 13
        Selection.Collapse (wdCollapseEnd)
    End If
    Selection.MoveRight (1)

End Sub

Sub insertar(fechaInicio As Date, fechaFinal As Date, capital As Double, datos() As String, dividirPorPeriodos As Boolean)

Dim periodos() As PeriodosDeIntereses
Dim i As Integer
Dim rng As Range
Dim tbl As Table
Dim total As Double

periodos = CalcularInteresesPorPeriodos(fechaInicio, fechaFinal, capital, datos)

For i = 0 To UBound(periodos)
    
    total = total + periodos(i).interesesPeriodo
    
Next i

Set rng = Selection.Range

If dividirPorPeriodos Then
       
    Set tbl = rng.Tables.Add(rng, UBound(periodos) + 3, 6)
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
            tbl.Rows(i).Cells(2).Range.text = periodos(i - 2).fechaInicioPeriodo
        End If
        
        If i = tbl.Rows.Count - 1 Then
            tbl.Rows(i).Cells(3).Range.text = fechaFinal
        Else
            tbl.Rows(i).Cells(3).Range.text = periodos(i - 2).fechaFinPeriodo
        End If
        tbl.Rows(i).Cells(4).Range.text = periodos(i - 2).dias
        tbl.Rows(i).Cells(5).Range.text = periodos(i - 2).tipo & "%"
        tbl.Rows(i).Cells(6).Range.text = FormatCurrency(periodos(i - 2).interesesPeriodo)
    
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


End Sub

Function CalcularInteresesPorPeriodos(fechaInicio As Date, fechaFin As Date, capital As Double, datos() As String) As PeriodosDeIntereses()

'Variable type privado donde se almacenará la información de rangos de fechas, sus tipos de interés, los días que contiene y los intereses devengados para ese periodo
Dim periodos() As PeriodosDeIntereses
'Variables para guardar la primera y la última fecha para la que hay información disponible
Dim tamañoDatos As Integer
Dim fechaFinDatos As Date
Dim fechaInicioDatos As Date
Dim indexPrimerPeriodo As Integer
Dim indexUltimoPeriodo As Integer
Dim totalPeriodos As Integer
Dim i As Integer
Dim indice As Integer

tamañoDatos = UBound(datos)
fechaFinDatos = datos(tamañoDatos)

indexPrimerPeriodo = LBound(datos)
fechaInicioDatos = datos(indexPrimerPeriodo)

'Valida que la fecha de inicio no sea superior a la fecha de fin

If fechaInicio > fechaFin Then

    MsgBox "La fecha de inicio (" & fechaInicio & ") no puede ser superior a la fecha de fin del cálculo (" & fechaFin & ")."

    Exit Function

End If

'Valida que la fecha de inicio no sea superior al primer periodo con datos disponibles

If CDate(fechaInicioDatos) > fechaInicio Then

    MsgBox "No existen datos anteriores a " & fechaInicioDatos & "."

    Exit Function

End If

 
'Para no cargar en periodos() todos los periodos disponibles, incluso los que no devengan intereses por estar fuera del rango de fechas utilizado, calculamos qué índice de parse() contiene el primer y el último periodo con intereses
Do While CDate(datos(indexPrimerPeriodo)) <= fechaInicio And indexPrimerPeriodo < tamañoDatos
    indexPrimerPeriodo = indexPrimerPeriodo + 2 ' avanzar de dos en dos, ya que parse() está organizado en pares fecha/tipo
Loop

indexPrimerPeriodo = indexPrimerPeriodo - 2

indexUltimoPeriodo = indexPrimerPeriodo
Do While CDate(datos(indexUltimoPeriodo)) <= fechaFin And indexUltimoPeriodo < tamañoDatos
    indexUltimoPeriodo = indexUltimoPeriodo + 2 ' avanzar de dos en dos, ya que parse() está organizado en pares fecha/tipo
Loop

indexUltimoPeriodo = indexUltimoPeriodo - 2

' Definir la cantidad de periodos a calcular

totalPeriodos = (indexUltimoPeriodo - indexPrimerPeriodo) / 2

' Si la fecha inicial es inferior al límite, pero la fecha final es superior al límite, hay que añadir un periodo adicional, que irá calculado al último tipo disponible, pero separado.
If fechaFin > fechaFinDatos And fechaInicio < fechaFinDatos Then
    totalPeriodos = totalPeriodos + 1
End If

ReDim periodos(totalPeriodos)


' recorre el array "parse", tomando los elementos pares (indice = 0, 2, 4...), que se corresponden con las fechas, y los asigna al array "periodos.fechaInicioPeriodo".
' a continuación toma los elementos impares (indice+1), correspondientes a los tipos de interés, y los asigna al array "periodos.tipo".
' acaba en "tamañoDatos - 1" para no incluir la fecha final (fin del último periodo sobre el que hay datos), que va en una variable aparte

i = 0

If fechaInicio > fechaFinDatos And fechaInicio > fechaFinDatos Then

    With periodos(i)
        .fechaInicioPeriodo = fechaInicio
        .tipo = datos(tamañoDatos - 1)
        .fechaFinPeriodo = fechaFin
        
    End With

Else

For indice = indexPrimerPeriodo To indexUltimoPeriodo Step 2


    With periodos(i)
        .fechaInicioPeriodo = datos(indice)
        .tipo = datos(indice + 1)
        
        If (indice + 2) < tamañoDatos Then

            .fechaFinPeriodo = CDate(datos(indice + 2)) - 1
        Else

            .fechaFinPeriodo = datos(indice + 2)

        End If
        
    End With
    
    
    i = i + 1

Next indice



'Como el loop anterior acaba siempre con un i = i + 1, si i = totalPeriodos, significa que hay un periodo adicional (fuera de rango de tipos disponibles)

If i = totalPeriodos Then

    With periodos(i)
        .fechaInicioPeriodo = fechaFinDatos + 1
        .tipo = datos(tamañoDatos - 1)
        
        .fechaFinPeriodo = fechaFin
        
    End With

End If

End If


For i = 0 To totalPeriodos

    With periodos(i)

        If i = 0 Then

            If fechaFin > .fechaFinPeriodo Then

                .dias = DateDiff("d", fechaInicio, .fechaFinPeriodo) + 1
            Else

                .dias = DateDiff("d", fechaInicio, fechaFin) + 1

            End If

        Else
                
            If fechaFin > .fechaFinPeriodo Then
    
                .dias = DateDiff("d", .fechaInicioPeriodo, .fechaFinPeriodo) + 1
            Else
    
                .dias = DateDiff("d", .fechaInicioPeriodo, fechaFin) + 1
    
            End If
    
                '.Dias = DateDiff("d", .fechaInicio, .fechaFinal) + 1
        End If

        .interesesPeriodo = capital * .tipo * .dias / (DateDiff("d", DateSerial(year(.fechaInicioPeriodo), 1, 1), DateSerial(year(.fechaInicioPeriodo), 12, 31)) + 1) / 100

    End With


Next i

CalcularInteresesPorPeriodos = periodos


End Function

Sub test()
Dim Intereses As String
Dim datos() As String
Dim a() As PeriodosDeIntereses

Intereses = "01/01/1995:9:01/01/1996:9:01/01/1997:7,5:01/01/1998:5,5:01/01/1999:4,25:01/01/2000:4,25:01/01/2001:5,5:01/01/2002:4,25:01/01/2003:4,25:01/01/2004:3,75:01/01/2005:4:01/01/2006:4:01/01/2007:5:01/01/2008:5,5:01/01/2009:5,5:01/04/2009:4:01/01/2010:4:01/01/2011:4:01/01/2012:4:01/01/2013:4:01/01/2014:4:01/01/2015:3,5:01/01/2016:3:01/01/2017:3:01/01/2018:3:01/01/2019:3:01/01/2020:3:01/01/2021:3:01/01/2022:3:01/01/2023:3,25:01/01/2024:3,25:31/12/2024"
datos = Split(Intereses, ":")
a = CalcularInteresesPorPeriodos(#1/1/2001#, #1/1/2003#, 10000, datos)

End Sub
