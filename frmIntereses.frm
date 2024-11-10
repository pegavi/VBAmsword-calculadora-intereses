VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIntereses 
   Caption         =   "Cálculo de Intereses"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   OleObjectBlob   =   "frmIntereses.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmIntereses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Dim Intereses() As String


Private Sub chbFinal_Click()

End Sub

Private Sub cmbCalcular_Click()
Dim parse() As String
Dim fechaInicial As Date
Dim fechaFinal As Date
Dim question As Integer

question = vbOK

If Me.cbxTipo.text <> "Personalizado" Then

    parse = Split(Mid(Intereses((Me.cbxTipo.ListIndex + 1)), InStr(Intereses((Me.cbxTipo.ListIndex + 1)), ":") + 1), ":")
    
Else

    parse = Split(Me.cbxDiaInicio.text + "/" + Me.cbxMesInicio.text + "/" + Me.cbxAnyInicio.text + ":" + CStr(CDbl(Me.txtTipoPersonal)) + ":" + Me.cbxDiaFin.text + "/" + Me.cbxMesFin.text + "/" + Me.cbxAnyFin.text, ":")

End If

fechaInicial = parse(0)
fechaFinal = parse(UBound(parse))

If DateSerial(Me.cbxAnyInicio, Me.cbxMesInicio, Me.cbxDiaInicio) < fechaInicial Then

    MsgBox "No existen datos anteriores a " & fechaInicial & " para este tipo de interés. Selecciona otra fecha inicial."

ElseIf DateSerial(Me.cbxAnyInicio, Me.cbxMesInicio, Me.cbxDiaInicio) > DateSerial(Me.cbxAnyFin, Me.cbxMesFin, Me.cbxDiaFin) Then

    MsgBox "La fecha de inicio (" & DateSerial(Me.cbxAnyInicio, Me.cbxMesInicio, Me.cbxDiaInicio) & ") no puede ser superior a la fecha de fin del cálculo (" & DateSerial(Me.cbxAnyFin, Me.cbxMesFin, Me.cbxDiaFin) & ")."

Else

    If DateSerial(Me.cbxAnyFin, Me.cbxMesFin, Me.cbxDiaFin) > fechaFinal Then

        question = MsgBox("La fecha final es superior a la última fecha de la que se disponen datos (" & fechaFinal & "). Los intereses a partir de esa fecha se calcularán conforme al último tipo de interés disponible. ¿Quieres seguir adelante?", vbOKCancel + vbQuestion)
        
        If question <> vbOK Then
        
            Exit Sub
        
        End If
    End If

    If IsNumeric(Me.txtCapital) Then

        Call insertar(DateSerial(Me.cbxAnyInicio, Me.cbxMesInicio, Me.cbxDiaInicio), DateSerial(Me.cbxAnyFin, Me.cbxMesFin, Me.cbxDiaFin), CDbl(Me.txtCapital), parse, Me.chbFinal.Value)
        Unload Me
    
    Else

        MsgBox "No has introducido una cantidad válida"
    
    End If


End If



End Sub

Private Sub cmbCancelar_Click()

Unload Me

End Sub
Private Sub cbxTipo_Change()
Dim cambio As Integer
Dim i As Integer

If Me.cbxTipo.text <> "Personalizado" Then

    Me.lblInicio.Top = 54
    Me.cbxDiaInicio.Top = 72
    Me.cbxMesInicio.Top = 72
    Me.cbxAnyInicio.Top = 72
    Me.lblFin.Top = 96
    Me.cbxDiaFin.Top = 114
    Me.cbxMesFin.Top = 114
    Me.cbxAnyFin.Top = 114
    Me.cmbFinHoy.Top = 114
    Me.cmbCalcular.Top = 138
    Me.cmbCancelar.Top = 138
    Me.chbFinal.Top = 168
    
    Me.chbFinal.Visible = True
    Me.txtTipoPersonal.Visible = False
    Me.lblIntPersonal.Visible = False
    Me.chbFinal.Enabled = True
    Me.lblPercent.Visible = False
    
    Me.Height = 222

Else

    For i = 0 To Me.Controls.Count - 1
    
        If Me.Controls(i).Tag <> "superior" Then
        
            Me.Controls(i).Top = Me.Controls(i).Top + 30
        
        End If
        
    Next i
    
    Me.Height = Me.Height + 10
    
    Me.chbFinal.Visible = False
    
    Me.lblIntPersonal.Left = 12
    Me.txtTipoPersonal.Left = 42
    Me.lblPercent.Left = 78
    
    Me.txtTipoPersonal.Visible = True
    Me.lblIntPersonal.Visible = True
    Me.lblPercent.Visible = True
    
    
End If



End Sub


Private Sub txtCapital_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Dim texto As String
    Dim posicion As Integer

    ' Permitir números (0-9), punto (.), coma (,) y tecla de retroceso
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 44 Or KeyAscii = 8 Then
        ' Obtener el texto actual del cuadro de texto
        texto = Me.txtCapital.text

        ' Insertar el carácter en la posición actual del cursor
        If KeyAscii <> 8 Then ' Si no es la tecla de retroceso
            posicion = Me.txtCapital.SelStart + 1
            texto = Left(texto, Me.txtCapital.SelStart) & Chr(KeyAscii) & Mid(texto, Me.txtCapital.SelStart + 1)
        Else
            If Me.txtCapital.SelLength > 0 Then
                texto = Left(texto, Me.txtCapital.SelStart) & Mid(texto, Me.txtCapital.SelStart + Me.txtCapital.SelLength + 1)
            ElseIf Me.txtCapital.SelStart > 0 Then
                texto = Left(texto, Me.txtCapital.SelStart - 1) & Mid(texto, Me.txtCapital.SelStart + 1)
            End If
        End If

        ' Reemplazar coma (,) por punto (.) para validación
       ' texto = Replace(texto, ",", ".")

        ' Validar si el resultado es un número válido
        If Not IsNumeric(texto) Then
            'MsgBox "Por favor, introduzca un número válido.", vbExclamation
            KeyAscii = 0 ' Cancelar la tecla
        End If
    Else
        'MsgBox "Por favor, introduzca solo números, punto o coma.", vbExclamation
        KeyAscii = 0 ' Cancelar la tecla
    End If




End Sub

Private Sub txtTipoPersonal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Dim texto As String
    Dim posicion As Integer

    ' Permitir números (0-9), punto (.), coma (,) y tecla de retroceso
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 44 Or KeyAscii = 8 Then
        ' Obtener el texto actual del cuadro de texto
        texto = Me.txtTipoPersonal.text

        ' Insertar el carácter en la posición actual del cursor
        If KeyAscii <> 8 Then ' Si no es la tecla de retroceso
            posicion = Me.txtTipoPersonal.SelStart + 1
            texto = Left(texto, Me.txtTipoPersonal.SelStart) & Chr(KeyAscii) & Mid(texto, Me.txtTipoPersonal.SelStart + 1)
        Else
            If Me.txtTipoPersonal.SelLength > 0 Then
                texto = Left(texto, Me.txtTipoPersonal.SelStart) & Mid(texto, Me.txtTipoPersonal.SelStart + Me.txtTipoPersonal.SelLength + 1)
            ElseIf Me.txtTipoPersonal.SelStart > 0 Then
                texto = Left(texto, Me.txtTipoPersonal.SelStart - 1) & Mid(texto, Me.txtTipoPersonal.SelStart + 1)
            End If
        End If

        ' Reemplazar coma (,) por punto (.) para validación
       ' texto = Replace(texto, ",", ".")

        ' Validar si el resultado es un número válido
        If Not IsNumeric(texto) Then
            'MsgBox "Por favor, introduzca un número válido.", vbExclamation
            KeyAscii = 0 ' Cancelar la tecla
        End If
    Else
        'MsgBox "Por favor, introduzca solo números, punto o coma.", vbExclamation
        KeyAscii = 0 ' Cancelar la tecla
    End If




End Sub

Private Sub UserForm_Activate()

' Medidas del Userform
' Height = 222.5
' Width = 220.5

Dim i As Integer
Dim datosGuardados() As String
Dim config As String
Dim pos As Integer

Me.Height = 222.5
Me.Width = 220.5

datosGuardados = Split(LeerStringDesdeTxt, ";")
config = datosGuardados(0)

ReDim Intereses(1 To UBound(datosGuardados))

For i = 1 To UBound(Intereses)
    Intereses(i) = datosGuardados(i)
Next i


Dim anno As Integer


For i = 1 To UBound(Intereses)

    pos = InStr(Intereses(i), ":")

    If Intereses(i) <> "" Then

        Me.cbxTipo.AddItem (Left(Intereses(i), pos - 1))

    End If

Next i

Me.cbxTipo.AddItem ("Personalizado")

Me.cbxTipo.ListIndex = 0


For i = 1 To 31

    Me.cbxDiaInicio.AddItem i
    Me.cbxDiaFin.AddItem i

Next i

For i = 1 To 12

    Me.cbxMesInicio.AddItem i
    Me.cbxMesFin.AddItem i

Next i
    
    
    pos = InStr(Intereses(1), ":")
For i = 0 To Year(Now) - CInt(Mid(Intereses(1), pos + 7, 4))

    anno = CInt(Mid(Intereses(1), pos + 7, 4)) + i
    
    Me.cbxAnyInicio.AddItem anno
    Me.cbxAnyFin.AddItem anno
    
    
Next i

If IsNumeric(Selection.Range.text) Then

    Me.txtCapital.Value = Selection.Range.text

End If

End Sub

Private Sub cbxAnyInicio_Change()

Call cbxMesInicio_Change

End Sub

Private Sub cbxAnyFin_Change()

Call cbxMesFin_Change

End Sub

Private Sub cbxMesFin_Change()

Dim mes As Long
Dim i As Integer
Dim anno As Integer
Dim day As Integer

day = Me.cbxDiaFin.Value
mes = Me.cbxMesFin.Value
anno = Me.cbxAnyFin.Value

Me.cbxDiaFin.Clear


If mes = 1 Or mes = 3 Or mes = 5 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12 Then

    For i = 1 To 31

        Me.cbxDiaFin.AddItem i
    
    Next i
    
    Me.cbxDiaFin.Value = day
    
ElseIf mes = 2 Then

    If ((anno - 2000) Mod 4) = 0 Then
    
        For i = 1 To 29

            Me.cbxDiaFin.AddItem i
        
        Next i
        
        If day > 29 Then
        
            Me.cbxDiaFin.Value = 29
            
        Else
            
            Me.cbxDiaFin.Value = day

        End If
          
    Else
    
    For i = 1 To 28

            Me.cbxDiaFin.AddItem i
        
        Next i
        
        If day > 28 Then
        
            Me.cbxDiaFin.Value = 28
            
        Else
            
            Me.cbxDiaFin.Value = day

        End If

            
    End If

Else

    For i = 1 To 30

        Me.cbxDiaFin.AddItem i
    
    Next i

        If day = 31 Then
        
            Me.cbxDiaFin.Value = 30
            
        Else
            
            Me.cbxDiaFin.Value = day

        End If


End If


End Sub

Private Sub cbxMesInicio_Change()

Dim mes As Long
Dim i As Integer
Dim anno As Integer
Dim day As Integer

day = Me.cbxDiaInicio.Value
mes = Me.cbxMesInicio.Value
anno = Me.cbxAnyInicio.Value

Me.cbxDiaInicio.Clear


If mes = 1 Or mes = 3 Or mes = 5 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12 Then

    For i = 1 To 31

        Me.cbxDiaInicio.AddItem i
    
    Next i
    
    Me.cbxDiaInicio.Value = day
    
ElseIf mes = 2 Then

    If ((anno - 2000) Mod 4) = 0 Then
    
        For i = 1 To 29

            Me.cbxDiaInicio.AddItem i
        
        Next i
        
        If day > 29 Then
        
            Me.cbxDiaInicio.Value = 29
            
        Else
            
            Me.cbxDiaInicio.Value = day

        End If
          
    Else
    
    For i = 1 To 28

            Me.cbxDiaInicio.AddItem i
        
        Next i
        
        If day > 28 Then
        
            Me.cbxDiaInicio.Value = 28
            
        Else
            
            Me.cbxDiaInicio.Value = day

        End If

            
    End If

Else

    For i = 1 To 30

        Me.cbxDiaInicio.AddItem i
    
    Next i

        If day = 31 Then
        
            Me.cbxDiaInicio.Value = 30
            
        Else
            
            Me.cbxDiaInicio.Value = day

        End If


End If


End Sub

Private Sub cmbFinHoy_Click()

Me.cbxAnyFin.Value = Year(Date)
Me.cbxMesFin.Value = month(Date)
Me.cbxDiaFin.Value = day(Date)

End Sub

Private Function LeerStringDesdeTxt()
    Dim filePath As String
    Dim fileNumber As Integer
    Dim textLine As String
    Dim contenidoCompleto As String
    
    ' Obtener la ruta del documento actual (Normal.dotm)
    filePath = ThisDocument.Path & "\configIntereses.txt"
    
    ' Comprobar si el archivo existe, y en caso contrario añadir lista de intereses por defecto
    ' La idea sería en un futuro crear un txt. con estos datos por defecto
    If Dir(filePath) = "" Then
        contenidoCompleto = " ;" & _
        "Interés Legal" & ":" & "01/01/1995:9:01/01/1996:9:01/01/1997:7,5:01/01/1998:5,5:01/01/1999:4,25:01/01/2000:4,25:01/01/2001:5,5:01/01/2002:4,25:01/01/2003:4,25:01/01/2004:3,75:01/01/2005:4:01/01/2006:4:01/01/2007:5:01/01/2008:5,5:01/01/2009:5,5:01/04/2009:4:01/01/2010:4:01/01/2011:4:01/01/2012:4:01/01/2013:4:01/01/2014:4:01/01/2015:3,5:01/01/2016:3:01/01/2017:3:01/01/2018:3:01/01/2019:3:01/01/2020:3:01/01/2021:3:01/01/2022:3:01/01/2023:3,25:01/01/2024:3,25:31/12/2024" & ";" & _
        "Procesal (legal +2 pts)" & ":" & "01/01/1995:11:01/01/1996:11:01/01/1997:9,5:01/01/1998:7,5:01/01/1999:6,25:01/01/2000:6,25:01/01/2001:7,5:01/01/2002:6,25:01/01/2003:6,25:01/01/2004:5,75:01/01/2005:6:01/01/2006:6:01/01/2007:7:01/01/2008:7,5:01/01/2009:7,5:01/04/2009:6:01/01/2010:6:01/01/2011:6:01/01/2012:6:01/01/2013:6:01/01/2014:6:01/01/2015:5,5:01/01/2016:5:01/01/2017:5:01/01/2018:5:01/01/2019:5:01/01/2020:5:01/01/2021:5:01/01/2022:5:01/01/2023:5,25:01/01/2024:5,25:31/12/2024" & ";" & _
        "Op. Comerciales" & ":" & "01/01/2003:9,85:01/07/2003:9,1:01/01/2004:9,02:01/07/2004:9,01:01/01/2005:9,09:01/07/2005:9,05:01/01/2006:9,25:01/07/2006:9,83:01/01/2007:10,58:01/07/2007:11,07:01/01/2008:11,2:01/07/2008:11,07:01/01/2009:9,5:01/07/2009:8:01/01/2010:8:01/07/2010:8:01/01/2011:8:01/07/2011:8,25:01/01/2012:8:01/07/2012:8:01/01/2013:7,75:24/02/2013:8,75:01/07/2013:8,5:01/01/2014:8,25:01/07/2014:8,15:01/01/2015:8,05:01/07/2015:8,05:01/01/2016:8,05:01/07/2016:8:01/01/2017:8:01/07/2017:8:01/01/2018:8:01/07/2018:8:01/01/2019:8:01/07/2019:8:01/01/2020:8:01/07/2020:8:01/01/2021:8:01/07/2021:8:01/01/2022:8:01/07/2022:8:01/01/2023:10,5:01/07/2023:12:01/01/2024:12,5:30/06/2024"
        
        LeerStringDesdeTxt = contenidoCompleto
        Exit Function
        
    End If
    
    ' Abrir el archivo para leer
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
       
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, textLine
        contenidoCompleto = contenidoCompleto & textLine & ";"
    Loop
    
    ' Cerrar el archivo
    Close #fileNumber
    
    'Eliminar el ultimo ";" que se añade, para evitar que posteriormente el split añada un elemento de más
    contenidoCompleto = Left(contenidoCompleto, Len(contenidoCompleto) - 1)
    
    LeerStringDesdeTxt = contenidoCompleto
    

End Function

Sub test()

Call frmIntereses

End Sub

