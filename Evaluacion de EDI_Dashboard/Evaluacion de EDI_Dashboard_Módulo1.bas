Attribute VB_Name = "Módulo1"
Option Explicit

' ============================================================================
' CÁLCULO EXACTO DE MESES (CRONOLÓGICOS O CORREGIDOS) SIN RETRASOS
' ============================================================================
Public Function CalcularMesesExactosEDI(ByVal fNac As Date, ByVal fAtencion As Date, ByVal semanasGest As Integer) As Double
    Dim anioA As Integer, mesA As Integer, diaA As Integer
    Dim anioN As Integer, mesN As Integer, diaN As Integer
    Dim d As Integer, m As Integer, a As Integer
    Dim semPrem As Integer
    Dim aCorr As Integer, mCorr As Integer, dCorr As Integer
    Dim totalDiasCron As Long, diasPrem As Long
    Dim aRest As Integer, restoDias As Long, mRest As Integer, dRest As Integer
    
    anioA = Year(fAtencion): mesA = Month(fAtencion): diaA = Day(fAtencion)
    anioN = Year(fNac): mesN = Month(fNac): diaN = Day(fNac)
    
    d = diaA - diaN
    m = mesA - mesN
    a = anioA - anioN
    
    If d < 0 Then
        d = d + 30
        m = m - 1
    End If
    If m < 0 Then
        m = m + 12
        a = a - 1
    End If
    
    If semanasGest = 0 Or semanasGest >= 37 Then
        semPrem = 0
    Else
        semPrem = 40 - semanasGest
    End If
    
    aCorr = a: mCorr = m: dCorr = d
    
    If semPrem > 0 And a < 2 Then
        totalDiasCron = (CLng(a) * 360) + (CLng(m) * 30) + d
        diasPrem = CLng(semPrem) * 7
        
        If totalDiasCron <= diasPrem Then
            aCorr = 0: mCorr = 0: dCorr = 0
        Else
            aRest = Int(diasPrem / 360)
            restoDias = diasPrem Mod 360
            mRest = Int(restoDias / 30)
            dRest = restoDias Mod 30
            
            dCorr = d - dRest
            mCorr = m - mRest
            aCorr = a - aRest
            
            If dCorr < 0 Then
                dCorr = dCorr + 30
                mCorr = mCorr - 1
            End If
            If mCorr < 0 Then
                mCorr = mCorr + 12
                aCorr = aCorr - 1
            End If
        End If
    End If
    
    If aCorr > 0 Then
        CalcularMesesExactosEDI = (aCorr * 12) + mCorr
    Else
        CalcularMesesExactosEDI = mCorr
    End If
End Function

Public Function ObtenerGrupoPorMeses(ByVal meses As Double) As Integer
    Dim m As Integer
    m = Int(meses)
    If m < 1 Then ObtenerGrupoPorMeses = 0: Exit Function
    If m = 1 Then ObtenerGrupoPorMeses = 1: Exit Function
    If m = 2 Then ObtenerGrupoPorMeses = 2: Exit Function
    If m = 3 Then ObtenerGrupoPorMeses = 3: Exit Function
    If m = 4 Then ObtenerGrupoPorMeses = 4: Exit Function
    If m >= 5 And m <= 6 Then ObtenerGrupoPorMeses = 5: Exit Function
    If m >= 7 And m <= 9 Then ObtenerGrupoPorMeses = 6: Exit Function
    If m >= 10 And m <= 12 Then ObtenerGrupoPorMeses = 7: Exit Function
    If m >= 13 And m <= 15 Then ObtenerGrupoPorMeses = 8: Exit Function
    If m >= 16 And m <= 18 Then ObtenerGrupoPorMeses = 9: Exit Function
    If m >= 19 And m <= 24 Then ObtenerGrupoPorMeses = 10: Exit Function
    If m >= 25 And m <= 30 Then ObtenerGrupoPorMeses = 11: Exit Function
    If m >= 31 And m <= 36 Then ObtenerGrupoPorMeses = 12: Exit Function
    If m >= 37 And m <= 48 Then ObtenerGrupoPorMeses = 13: Exit Function
    If m >= 49 And m <= 60 Then ObtenerGrupoPorMeses = 14: Exit Function
    If m >= 61 And m <= 71 Then ObtenerGrupoPorMeses = 15: Exit Function
    ObtenerGrupoPorMeses = 99
End Function

Private Function ExtraerNumeroGrupo(ByVal texto As String) As Integer
    Dim i As Integer, ch As String, digitos As String
    digitos = ""
    For i = 1 To Len(texto)
        ch = Mid(texto, i, 1)
        If ch >= "0" And ch <= "9" Then
            digitos = digitos & ch
        ElseIf digitos <> "" Then
            Exit For
        End If
    Next i
    If digitos = "" Then
        ExtraerNumeroGrupo = -1
    Else
        ExtraerNumeroGrupo = CInt(digitos)
    End If
End Function

Private Function ObtenerColorPastelAreaEDI(ByVal tipoSeccion As String, ByVal subArea As String) As Long
    Dim secL As String, areaL As String
    secL = LCase(Trim(tipoSeccion))
    areaL = LCase(Trim(subArea))
    
    If InStr(secL, "desarrollo") > 0 Or InStr(secL, "área") > 0 Then
        If InStr(areaL, "gruesa") > 0 Then
            ObtenerColorPastelAreaEDI = RGB(228, 245, 224)
        ElseIf InStr(areaL, "fina") > 0 Then
            ObtenerColorPastelAreaEDI = RGB(222, 238, 250)
        ElseIf InStr(areaL, "lenguaje") > 0 Then
            ObtenerColorPastelAreaEDI = RGB(252, 228, 236)
        ElseIf InStr(areaL, "social") > 0 Then
            ObtenerColorPastelAreaEDI = RGB(255, 242, 217)
        ElseIf InStr(areaL, "conocimiento") > 0 Then
            ObtenerColorPastelAreaEDI = RGB(237, 228, 247)
        Else
            ObtenerColorPastelAreaEDI = RGB(245, 247, 250)
        End If
    ElseIf InStr(secL, "alarma") > 0 Or InStr(secL, "neurol") > 0 Then
        ObtenerColorPastelAreaEDI = RGB(242, 244, 247)
    Else
        ObtenerColorPastelAreaEDI = RGB(250, 248, 240)
    End If
End Function

' ============================================================================
' 1. BOTÓN "1. CARGAR PREGUNTAS (AUTO)"
' ============================================================================
Public Sub CargarPreguntasAutomaticas()
    Dim dash As Worksheet, hojaCalculo As Worksheet, base As Worksheet
    Dim nombreNino As String, strGest As String
    Dim fAtencion As Date, fNac As Date
    Dim valorGestNum As Integer, mesesNum As Double, numGrupoInt As Integer
    Dim codigoGrupoStr As String, codigoGrupoAntStr As String, infoEdadStr As String
    Dim ultimaFilaBase As Long, i As Long
    Dim tablaFinal() As Variant
    Dim totalFilas As Long, idx As Long
    Dim numBase As Integer, grupoBase As String
    Dim tipoSeccion As String, subArea As String, modVal As String, modSimbolo As String, textoPreg As String
    Dim rangoTabla As Range
    Dim rCel As Range
    Const FILA_INI As Long = 24
    Const FILA_FIN As Long = 75
    
    On Error Resume Next
    Set dash = ThisWorkbook.Sheets("Dashboard")
    Set hojaCalculo = ThisWorkbook.Sheets("Calculo")
    Set base = ThisWorkbook.Sheets("Base_Preguntas")
    On Error GoTo 0
    
    If dash Is Nothing Or hojaCalculo Is Nothing Or base Is Nothing Then
        MsgBox "Verifica que existan las pestañas 'Dashboard', 'Calculo' y 'Base_Preguntas'.", vbCritical, "Error de Hojas"
        Exit Sub
    End If
    
    nombreNino = Trim(CStr(dash.Range("B6").MergeArea.Cells(1, 1).Value))
    
    If Not IsDate(dash.Range("F5").MergeArea.Cells(1, 1).Value) Or dash.Range("F5").MergeArea.Cells(1, 1).Value = "" Then
        MsgBox "Ingresa una Fecha de Atención válida en F5.", vbExclamation, "Error en Fecha"
        Exit Sub
    End If
    If Not IsDate(dash.Range("F6").MergeArea.Cells(1, 1).Value) Or dash.Range("F6").MergeArea.Cells(1, 1).Value = "" Then
        MsgBox "Ingresa una Fecha de Nacimiento válida en F6.", vbExclamation, "Error en Fecha"
        Exit Sub
    End If
    
    fAtencion = CDate(dash.Range("F5").MergeArea.Cells(1, 1).Value)
    fNac = CDate(dash.Range("F6").MergeArea.Cells(1, 1).Value)
    
    If fNac > fAtencion Then
        MsgBox "La Fecha de Nacimiento no puede ser mayor que la Fecha de Atención.", vbCritical, "Fecha Incorrecta"
        Exit Sub
    End If
    If nombreNino = "" Then
        MsgBox "Por favor ingresa el Nombre del Niño en B6.", vbExclamation, "Campo Requerido"
        Exit Sub
    End If
    
    strGest = Trim(CStr(dash.Range("F7").MergeArea.Cells(1, 1).Value))
    If strGest <> "" Then
        If Not IsNumeric(strGest) Then
            MsgBox "Edad gestacional no válida.", vbExclamation
            Exit Sub
        End If
        Dim nG As Double
        nG = CDbl(strGest)
        If Int(nG) <> nG Or (nG <> 0 And (nG < 21 Or nG > 40)) Then
            MsgBox "Edad gestacional no válida:" & vbCrLf & _
                   "- Ingrese un número entero entre 21 y 40 semanas (o deje vacío / 0).", vbExclamation, "Edad Gestacional Inválida"
            Exit Sub
        End If
        valorGestNum = CInt(nG)
    Else
        valorGestNum = 0
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    hojaCalculo.Range("D8").Value = Year(fAtencion)
    hojaCalculo.Range("E8").Value = Month(fAtencion)
    hojaCalculo.Range("F8").Value = Day(fAtencion)
    
    hojaCalculo.Range("D9").Value = Year(fNac)
    hojaCalculo.Range("E9").Value = Month(fNac)
    hojaCalculo.Range("F9").Value = Day(fNac)
    
    hojaCalculo.Range("G10").Value = valorGestNum
    hojaCalculo.Range("D13").Value = valorGestNum
    hojaCalculo.Calculate
    
    mesesNum = CalcularMesesExactosEDI(fNac, fAtencion, valorGestNum)
    numGrupoInt = ObtenerGrupoPorMeses(mesesNum)
    
    If mesesNum < 1 Or numGrupoInt < 1 Or numGrupoInt > 15 Then
        For Each rCel In dash.Range("A" & FILA_INI & ":H" & FILA_FIN)
            If rCel.MergeCells Then rCel.MergeArea.UnMerge
        Next rCel
        dash.Range("A" & FILA_INI & ":H" & FILA_FIN).ClearContents
        dash.Range("A" & FILA_INI & ":H" & FILA_FIN).Interior.ColorIndex = xlNone
        dash.Range("A" & FILA_INI & ":H" & FILA_FIN).Borders.LineStyle = xlNone
        dash.Range("F" & FILA_INI & ":F" & FILA_FIN).Validation.Delete
        
        Dim motivo As String
        motivo = IIf(mesesNum < 1, "Menor a 1 mes", "Mayor a 71 meses")
        dash.Range("B15").MergeArea.Cells(1, 1).Value = "No aplica (" & Format(mesesNum, "0.0") & " meses - " & motivo & ")"
        With dash.Range("C20").MergeArea.Cells(1, 1)
            .Value = "NO APLICA EVALUACIÓN EDI (" & UCase(motivo) & ")"
            .Interior.Color = RGB(232, 236, 239)
            .Font.Color = RGB(73, 80, 87)
        End With
        
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "La edad calculada es de " & Format(mesesNum, "0.0") & " meses." & vbCrLf & _
               "La prueba EDI se aplica únicamente desde 1 mes cumplido (Grupo 01) hasta los 71 meses (Grupo 15).", vbInformation, "Fuera de Rango EDI"
        Exit Sub
    End If
    
    codigoGrupoStr = Format(numGrupoInt, "00")
    codigoGrupoAntStr = Format(numGrupoInt - 1, "00")
    infoEdadStr = "Grupo " & codigoGrupoStr & " (" & Format(mesesNum, "0.0") & " meses)"
    dash.Range("B15").MergeArea.Cells(1, 1).Value = infoEdadStr
    
    ultimaFilaBase = base.Cells(base.Rows.Count, 1).End(xlUp).Row
    ReDim tablaFinal(1 To 60, 1 To 8)
    totalFilas = 0
    
    ' 1. Preguntas principales del grupo actual
    For i = 2 To ultimaFilaBase
        numBase = ExtraerNumeroGrupo(CStr(base.Cells(i, 2).Value))
        If numBase > 0 Then
            grupoBase = Format(numBase, "00")
            tipoSeccion = Trim(CStr(base.Cells(i, 4).Value))
            subArea = Trim(CStr(base.Cells(i, 5).Value))
            modVal = Trim(CStr(base.Cells(i, 7).Value))
            textoPreg = CStr(base.Cells(i, 8).Value)
            modSimbolo = IIf(modVal = "Sí" Or modVal = "SI" Or modVal = "¿?", "¿?", "**")
            
            If grupoBase = codigoGrupoStr Then
                If InStr(tipoSeccion, "Alarma") > 0 And numGrupoInt = 1 Then GoTo Siguiente1
                If InStr(subArea, "Conocimiento") > 0 And numGrupoInt < 13 Then GoTo Siguiente1
                If InStr(tipoSeccion, "Riesgo") = 0 And InStr(tipoSeccion, "Alerta") = 0 Then
                    totalFilas = totalFilas + 1
                    tablaFinal(totalFilas, 1) = base.Cells(i, 1).Value
                    tablaFinal(totalFilas, 2) = tipoSeccion
                    tablaFinal(totalFilas, 3) = subArea
                    tablaFinal(totalFilas, 4) = textoPreg
                    tablaFinal(totalFilas, 5) = modSimbolo
                    tablaFinal(totalFilas, 6) = ""
                    tablaFinal(totalFilas, 7) = "PENDIENTE"
                    tablaFinal(totalFilas, 8) = "-"
                End If
            End If
        End If
Siguiente1:
    Next i
    
    ' 2. Preguntas de Desarrollo del Grupo Anterior (si grupo >= 2)
    If numGrupoInt >= 2 Then
        For i = 2 To ultimaFilaBase
            numBase = ExtraerNumeroGrupo(CStr(base.Cells(i, 2).Value))
            If numBase > 0 Then
                grupoBase = Format(numBase, "00")
                tipoSeccion = Trim(CStr(base.Cells(i, 4).Value))
                subArea = Trim(CStr(base.Cells(i, 5).Value))
                modVal = Trim(CStr(base.Cells(i, 7).Value))
                textoPreg = CStr(base.Cells(i, 8).Value)
                modSimbolo = IIf(modVal = "Sí" Or modVal = "SI" Or modVal = "¿?", "¿?", "**")
                
                If grupoBase = codigoGrupoAntStr Then
                    If InStr(tipoSeccion, "Desarrollo") > 0 Or InStr(tipoSeccion, "Área") > 0 Then
                        totalFilas = totalFilas + 1
                        tablaFinal(totalFilas, 1) = CStr(base.Cells(i, 1).Value) & "_ANT"
                        tablaFinal(totalFilas, 2) = "Área Desarrollo (Grupo " & codigoGrupoAntStr & " Anterior)"
                        tablaFinal(totalFilas, 3) = subArea
                        tablaFinal(totalFilas, 4) = "[GRUPO ANTERIOR] " & textoPreg
                        tablaFinal(totalFilas, 5) = modSimbolo
                        tablaFinal(totalFilas, 6) = ""
                        tablaFinal(totalFilas, 7) = "OPCIONAL / PENDIENTE"
                        tablaFinal(totalFilas, 8) = "-"
                    End If
                End If
            End If
        Next i
    End If
    
    ' 3. Factores de Riesgo Biológico y Señales de Alerta
    For i = 2 To ultimaFilaBase
        numBase = ExtraerNumeroGrupo(CStr(base.Cells(i, 2).Value))
        If numBase > 0 Then
            grupoBase = Format(numBase, "00")
            tipoSeccion = Trim(CStr(base.Cells(i, 4).Value))
            subArea = Trim(CStr(base.Cells(i, 5).Value))
            modVal = Trim(CStr(base.Cells(i, 7).Value))
            textoPreg = CStr(base.Cells(i, 8).Value)
            modSimbolo = IIf(modVal = "Sí" Or modVal = "SI" Or modVal = "¿?", "¿?", "**")
            
            If grupoBase = codigoGrupoStr Then
                If InStr(tipoSeccion, "Riesgo") > 0 Or InStr(tipoSeccion, "Alerta") > 0 Then
                    totalFilas = totalFilas + 1
                    tablaFinal(totalFilas, 1) = base.Cells(i, 1).Value
                    If numGrupoInt <= 5 Then
                        tablaFinal(totalFilas, 2) = tipoSeccion
                        tablaFinal(totalFilas, 4) = textoPreg
                        tablaFinal(totalFilas, 7) = "PENDIENTE"
                    Else
                        tablaFinal(totalFilas, 2) = tipoSeccion & " (Opcional)"
                        tablaFinal(totalFilas, 4) = "[OPCIONAL - NO MODIFICA GLOBAL] " & textoPreg
                        tablaFinal(totalFilas, 7) = "OPCIONAL / PENDIENTE"
                    End If
                    tablaFinal(totalFilas, 3) = subArea
                    tablaFinal(totalFilas, 5) = modSimbolo
                    tablaFinal(totalFilas, 6) = ""
                    tablaFinal(totalFilas, 8) = "-"
                End If
            End If
        End If
    Next i
    
    If totalFilas = 0 Then
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "No se encontraron preguntas para el " & infoEdadStr & " en 'Base_Preguntas'.", vbExclamation
        Exit Sub
    End If
    
    For Each rCel In dash.Range("A" & FILA_INI & ":H" & FILA_FIN)
        If rCel.MergeCells Then rCel.MergeArea.UnMerge
    Next rCel
    
    dash.Range("A" & FILA_INI & ":H" & FILA_FIN).ClearContents
    dash.Range("A" & FILA_INI & ":H" & FILA_FIN).Interior.ColorIndex = xlNone
    dash.Range("A" & FILA_INI & ":H" & FILA_FIN).Borders.LineStyle = xlNone
    dash.Range("F" & FILA_INI & ":F" & FILA_FIN).Validation.Delete
    dash.Rows(FILA_INI & ":" & FILA_FIN).RowHeight = 20
    
    Set rangoTabla = dash.Range(dash.Cells(FILA_INI, 1), dash.Cells(FILA_INI - 1 + totalFilas, 8))
    For idx = 1 To totalFilas
        Dim c As Integer
        For c = 1 To 8
            dash.Cells(FILA_INI - 1 + idx, c).Value = tablaFinal(idx, c)
        Next c
        dash.Range(dash.Cells(FILA_INI - 1 + idx, 1), dash.Cells(FILA_INI - 1 + idx, 5)).Interior.Color = _
            ObtenerColorPastelAreaEDI(CStr(tablaFinal(idx, 2)), CStr(tablaFinal(idx, 3)))
    Next idx
    
    With rangoTabla
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(205, 210, 215)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    
    dash.Range(dash.Cells(FILA_INI, 1), dash.Cells(FILA_INI - 1 + totalFilas, 1)).HorizontalAlignment = xlCenter
    dash.Range(dash.Cells(FILA_INI, 5), dash.Cells(FILA_INI - 1 + totalFilas, 8)).HorizontalAlignment = xlCenter
    dash.Range(dash.Cells(FILA_INI, 7), dash.Cells(FILA_INI - 1 + totalFilas, 8)).Font.Bold = True
    dash.Range(dash.Cells(FILA_INI, 7), dash.Cells(FILA_INI - 1 + totalFilas, 8)).Font.Color = RGB(0, 0, 0)
    
    Dim fIni As Long, fFin As Long
    fIni = FILA_INI
    Do While fIni <= FILA_INI - 1 + totalFilas
        Dim secIni As String, areaIni As String
        Dim esDesarrollo As Boolean
        secIni = Trim(CStr(dash.Cells(fIni, 2).Value))
        areaIni = Trim(CStr(dash.Cells(fIni, 3).Value))
        esDesarrollo = (InStr(secIni, "Desarrollo") > 0 Or InStr(secIni, "Área") > 0)
        
        fFin = fIni
        If esDesarrollo Then
            Do While fFin < FILA_INI - 1 + totalFilas And _
                     Trim(CStr(dash.Cells(fFin + 1, 2).Value)) = secIni And _
                     Trim(CStr(dash.Cells(fFin + 1, 3).Value)) = areaIni
                fFin = fFin + 1
            Loop
        Else
            Do While fFin < FILA_INI - 1 + totalFilas And _
                     Trim(CStr(dash.Cells(fFin + 1, 2).Value)) = secIni
                fFin = fFin + 1
            Loop
        End If
        
        If fFin > fIni Then
            With dash.Range(dash.Cells(fIni, 8), dash.Cells(fFin, 8))
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        End If
        fIni = fFin + 1
    Loop
    
    rangoTabla.EntireRow.AutoFit
    
    With dash.Range(dash.Cells(FILA_INI, 6), dash.Cells(FILA_INI - 1 + totalFilas, 6)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="SI,NO"
    End With
    
    With dash.Range("C20").MergeArea.Cells(1, 1)
        .Value = "PENDIENTE DE RESPUESTAS"
        .Interior.Color = RGB(232, 236, 239)
        .Font.Color = RGB(0, 0, 0)
    End With
    
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Preguntas cargadas con sus colores por área correctamente para " & infoEdadStr, vbInformation, "Carga Exitosa"
End Sub

' ============================================================================
' 2. BOTÓN "2. CALCULAR DIAGNÓSTICO"
' ============================================================================
Public Sub CalcularDiagnosticoEDI()
    Dim dash As Worksheet, hojaCalculo As Worksheet
    Dim mesesNum As Double, numGrupoInt As Integer
    Dim i As Long, respondidas As Integer
    Dim cod As String, tipoSeccion As String, subArea As String, resp As String
    Dim celdaEstado As Range, celdaResArea As Range
    Dim frbAmarillos As Integer, aleAmarillos As Integer, almRojos As Integer, enRojos As Integer
    Dim frbResp As Integer, aleResp As Integer, almResp As Integer, enResp As Integer
    Const FILA_INI As Long = 24
    Const FILA_FIN As Long = 75
    
    Dim actTotal As Object, actSi As Object, actNo As Object
    Dim antTotal As Object, antSi As Object, antNo As Object
    Dim resAreaFinal As Object
    
    Set dash = ThisWorkbook.Sheets("Dashboard")
    Set hojaCalculo = ThisWorkbook.Sheets("Calculo")
    
    mesesNum = Val(hojaCalculo.Range("F18").Value)
    numGrupoInt = ObtenerGrupoPorMeses(mesesNum)
    
    If mesesNum < 1 Or numGrupoInt < 1 Or numGrupoInt > 15 Then
        MsgBox "No se puede calcular diagnóstico: la edad está fuera del rango EDI (1 a 71 meses).", vbExclamation
        Exit Sub
    End If
    
    Set actTotal = CreateObject("Scripting.Dictionary")
    Set actSi = CreateObject("Scripting.Dictionary")
    Set actNo = CreateObject("Scripting.Dictionary")
    Set antTotal = CreateObject("Scripting.Dictionary")
    Set antSi = CreateObject("Scripting.Dictionary")
    Set antNo = CreateObject("Scripting.Dictionary")
    Set resAreaFinal = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For i = FILA_INI To FILA_FIN
        cod = Trim(CStr(dash.Cells(i, 1).Value))
        If cod = "" Then GoTo SiguienteFila
        
        tipoSeccion = Trim(CStr(dash.Cells(i, 2).Value))
        subArea = Trim(CStr(dash.Cells(i, 3).Value))
        resp = UCase(Trim(CStr(dash.Cells(i, 6).Value)))
        Set celdaEstado = dash.Cells(i, 7)
        
        If InStr(cod, "_ANT") > 0 Then
            If Not antTotal.Exists(subArea) Then
                antTotal(subArea) = 0: antSi(subArea) = 0: antNo(subArea) = 0
            End If
            antTotal(subArea) = antTotal(subArea) + 1
        ElseIf InStr(tipoSeccion, "Desarrollo") > 0 Or InStr(tipoSeccion, "Área") > 0 Then
            If Not actTotal.Exists(subArea) Then
                actTotal(subArea) = 0: actSi(subArea) = 0: actNo(subArea) = 0
            End If
            actTotal(subArea) = actTotal(subArea) + 1
        End If
        
        If resp = "" Then
            If InStr(cod, "_ANT") > 0 Or (numGrupoInt >= 6 And (InStr(tipoSeccion, "Riesgo") > 0 Or InStr(tipoSeccion, "Alerta") > 0)) Then
                celdaEstado.Value = "OPCIONAL / PENDIENTE"
            Else
                celdaEstado.Value = "PENDIENTE"
            End If
            celdaEstado.Interior.Color = RGB(232, 236, 239)
            celdaEstado.Font.Color = RGB(0, 0, 0)
            GoTo SiguienteFila
        End If
        
        respondidas = respondidas + 1
        
        If InStr(cod, "_ANT") > 0 Then
            If resp = "SI" Then
                antSi(subArea) = antSi(subArea) + 1
                celdaEstado.Value = "AMARILLO"
                celdaEstado.Interior.Color = RGB(255, 243, 205)
                celdaEstado.Font.Color = RGB(133, 100, 4)
            Else
                antNo(subArea) = antNo(subArea) + 1
                celdaEstado.Value = "ROJO"
                celdaEstado.Interior.Color = RGB(248, 215, 218)
                celdaEstado.Font.Color = RGB(114, 28, 36)
            End If
        ElseIf InStr(tipoSeccion, "Desarrollo") > 0 Or InStr(tipoSeccion, "Área") > 0 Then
            If resp = "SI" Then
                actSi(subArea) = actSi(subArea) + 1
                celdaEstado.Value = "VERDE"
                celdaEstado.Interior.Color = RGB(212, 237, 218)
                celdaEstado.Font.Color = RGB(21, 87, 36)
            Else
                actNo(subArea) = actNo(subArea) + 1
                If numGrupoInt = 1 Then
                    celdaEstado.Value = "ROJO"
                    celdaEstado.Interior.Color = RGB(248, 215, 218)
                    celdaEstado.Font.Color = RGB(114, 28, 36)
                Else
                    celdaEstado.Value = "AMARILLO"
                    celdaEstado.Interior.Color = RGB(255, 243, 205)
                    celdaEstado.Font.Color = RGB(133, 100, 4)
                End If
            End If
        ElseIf InStr(tipoSeccion, "Neurol") > 0 Or InStr(tipoSeccion, "Alarma") > 0 Then
            If InStr(tipoSeccion, "Neurol") > 0 Then enResp = enResp + 1 Else almResp = almResp + 1
            If resp = "SI" Then
                celdaEstado.Value = "ROJO"
                celdaEstado.Interior.Color = RGB(248, 215, 218)
                celdaEstado.Font.Color = RGB(114, 28, 36)
                If InStr(tipoSeccion, "Neurol") > 0 Then enRojos = enRojos + 1 Else almRojos = almRojos + 1
            Else
                celdaEstado.Value = "VERDE"
                celdaEstado.Interior.Color = RGB(212, 237, 218)
                celdaEstado.Font.Color = RGB(21, 87, 36)
            End If
        ElseIf InStr(tipoSeccion, "Riesgo") > 0 Or InStr(tipoSeccion, "Alerta") > 0 Then
            If InStr(tipoSeccion, "Riesgo") > 0 Then frbResp = frbResp + 1 Else aleResp = aleResp + 1
            If resp = "SI" Then
                celdaEstado.Value = "AMARILLO"
                celdaEstado.Interior.Color = RGB(255, 243, 205)
                celdaEstado.Font.Color = RGB(133, 100, 4)
                If InStr(tipoSeccion, "Riesgo") > 0 Then frbAmarillos = frbAmarillos + 1 Else aleAmarillos = aleAmarillos + 1
            Else
                celdaEstado.Value = "VERDE"
                celdaEstado.Interior.Color = RGB(212, 237, 218)
                celdaEstado.Font.Color = RGB(21, 87, 36)
            End If
        End If
SiguienteFila:
    Next i
    
    If respondidas = 0 Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "Selecciona al menos una respuesta (SI/NO) antes de calcular.", vbExclamation
        Exit Sub
    End If
    
    ' Calificación por Área del Desarrollo (incluye regla oficial Grupo 1: 2+ Verde = VERDE, 1 o 0 Verde = ROJO)
    Dim kArea As Variant
    Dim conteoAreasVerde As Integer, conteoAreasAmarillo As Integer, conteoAreasRojo As Integer
    
    For Each kArea In actTotal.Keys
        If (actSi(kArea) + actNo(kArea)) = 0 Then
            resAreaFinal(kArea) = "PENDIENTE"
        ElseIf numGrupoInt = 1 Then
            ' Regla Grupo 1 (1 mes): 2+ obtenidas en Verde = VERDE; 1 o ninguna en Verde = ROJO
            If actSi(kArea) >= 2 Then
                resAreaFinal(kArea) = "VERDE"
                conteoAreasVerde = conteoAreasVerde + 1
            Else
                resAreaFinal(kArea) = "ROJO"
                conteoAreasRojo = conteoAreasRojo + 1
            End If
        Else
            ' Regla Grupos 2 a 15
            If actSi(kArea) >= 2 Then
                resAreaFinal(kArea) = "VERDE"
                conteoAreasVerde = conteoAreasVerde + 1
            ElseIf actSi(kArea) = 1 Then
                resAreaFinal(kArea) = "AMARILLO"
                conteoAreasAmarillo = conteoAreasAmarillo + 1
            Else
                If antTotal.Exists(kArea) And (antSi(kArea) + antNo(kArea)) > 0 Then
                    If antSi(kArea) >= 2 Then
                        resAreaFinal(kArea) = "AMARILLO"
                        conteoAreasAmarillo = conteoAreasAmarillo + 1
                    Else
                        resAreaFinal(kArea) = "ROJO"
                        conteoAreasRojo = conteoAreasRojo + 1
                    End If
                Else
                    resAreaFinal(kArea) = "AMARILLO"
                    conteoAreasAmarillo = conteoAreasAmarillo + 1
                End If
            End If
        End If
    Next kArea
    
    For i = FILA_INI To FILA_FIN
        cod = Trim(CStr(dash.Cells(i, 1).Value))
        If cod = "" Then GoTo SiguienteColH
        
        tipoSeccion = Trim(CStr(dash.Cells(i, 2).Value))
        subArea = Trim(CStr(dash.Cells(i, 3).Value))
        Set celdaResArea = dash.Cells(i, 8)
        
        If InStr(cod, "_ANT") > 0 Then
            If (antSi(subArea) + antNo(subArea)) = 0 Then
                celdaResArea.Value = "NO APLICA"
                celdaResArea.Interior.Color = RGB(232, 236, 239)
                celdaResArea.Font.Color = RGB(73, 80, 87)
            ElseIf antSi(subArea) >= 2 Then
                celdaResArea.Value = "AMARILLO"
                celdaResArea.Interior.Color = RGB(255, 243, 205)
                celdaResArea.Font.Color = RGB(133, 100, 4)
            Else
                celdaResArea.Value = "ROJO"
                celdaResArea.Interior.Color = RGB(248, 215, 218)
                celdaResArea.Font.Color = RGB(114, 28, 36)
            End If
        ElseIf InStr(tipoSeccion, "Desarrollo") > 0 Or InStr(tipoSeccion, "Área") > 0 Then
            Dim estA As String
            estA = resAreaFinal(subArea)
            celdaResArea.Value = estA
            If estA = "VERDE" Then
                celdaResArea.Interior.Color = RGB(212, 237, 218)
                celdaResArea.Font.Color = RGB(21, 87, 36)
            ElseIf estA = "AMARILLO" Then
                celdaResArea.Interior.Color = RGB(255, 243, 205)
                celdaResArea.Font.Color = RGB(133, 100, 4)
            ElseIf estA = "ROJO" Then
                celdaResArea.Interior.Color = RGB(248, 215, 218)
                celdaResArea.Font.Color = RGB(114, 28, 36)
            Else
                celdaResArea.Interior.Color = RGB(232, 236, 239)
                celdaResArea.Font.Color = RGB(0, 0, 0)
            End If
        ElseIf InStr(tipoSeccion, "Alarma") > 0 Then
            If almResp = 0 Then
                celdaResArea.Value = "PENDIENTE"
                celdaResArea.Interior.Color = RGB(232, 236, 239)
                celdaResArea.Font.Color = RGB(0, 0, 0)
            ElseIf almRojos >= 1 Then
                celdaResArea.Value = "ROJO"
                celdaResArea.Interior.Color = RGB(248, 215, 218)
                celdaResArea.Font.Color = RGB(114, 28, 36)
            Else
                celdaResArea.Value = "VERDE"
                celdaResArea.Interior.Color = RGB(212, 237, 218)
                celdaResArea.Font.Color = RGB(21, 87, 36)
            End If
        ElseIf InStr(tipoSeccion, "Neurol") > 0 Then
            If enResp = 0 Then
                celdaResArea.Value = "PENDIENTE"
                celdaResArea.Interior.Color = RGB(232, 236, 239)
                celdaResArea.Font.Color = RGB(0, 0, 0)
            ElseIf enRojos >= 1 Then
                celdaResArea.Value = "ROJO"
                celdaResArea.Interior.Color = RGB(248, 215, 218)
                celdaResArea.Font.Color = RGB(114, 28, 36)
            Else
                celdaResArea.Value = "VERDE"
                celdaResArea.Interior.Color = RGB(212, 237, 218)
                celdaResArea.Font.Color = RGB(21, 87, 36)
            End If
        ElseIf InStr(tipoSeccion, "Riesgo") > 0 Then
            If frbResp = 0 Then
                celdaResArea.Value = IIf(numGrupoInt <= 5, "PENDIENTE", "NO APLICA")
                celdaResArea.Interior.Color = RGB(232, 236, 239)
                celdaResArea.Font.Color = RGB(0, 0, 0)
            ElseIf frbAmarillos >= 1 Then
                celdaResArea.Value = "AMARILLO"
                celdaResArea.Interior.Color = RGB(255, 243, 205)
                celdaResArea.Font.Color = RGB(133, 100, 4)
            Else
                celdaResArea.Value = "VERDE"
                celdaResArea.Interior.Color = RGB(212, 237, 218)
                celdaResArea.Font.Color = RGB(21, 87, 36)
            End If
        ElseIf InStr(tipoSeccion, "Alerta") > 0 Then
            If aleResp = 0 Then
                celdaResArea.Value = IIf(numGrupoInt <= 5, "PENDIENTE", "NO APLICA")
                celdaResArea.Interior.Color = RGB(232, 236, 239)
                celdaResArea.Font.Color = RGB(0, 0, 0)
            ElseIf aleAmarillos >= 1 Then
                celdaResArea.Value = "AMARILLO"
                celdaResArea.Interior.Color = RGB(255, 243, 205)
                celdaResArea.Font.Color = RGB(133, 100, 4)
            Else
                celdaResArea.Value = "VERDE"
                celdaResArea.Interior.Color = RGB(212, 237, 218)
                celdaResArea.Font.Color = RGB(21, 87, 36)
            End If
        End If
SiguienteColH:
    Next i
    
    Dim frbEfectivos As Integer, aleEfectivos As Integer
    If numGrupoInt <= 5 Then
        frbEfectivos = frbAmarillos
        aleEfectivos = aleAmarillos
    Else
        frbEfectivos = 0
        aleEfectivos = 0
    End If
    
    With dash.Range("C20").MergeArea.Cells(1, 1)
        If conteoAreasRojo >= 1 Or _
           conteoAreasAmarillo >= 2 Or _
           (conteoAreasAmarillo = 1 And (frbEfectivos > 0 Or aleEfectivos > 0)) Or _
           almRojos >= 1 Or _
           enRojos >= 1 Then
            .Value = "RIESGO DE RETRASO EN EL DESARROLLO (ROJO)"
            .Interior.Color = RGB(248, 215, 218)
            .Font.Color = RGB(114, 28, 36)
        ElseIf conteoAreasAmarillo = 1 Or _
               aleEfectivos >= 2 Or _
               frbEfectivos >= 2 Or _
               (aleEfectivos >= 1 And frbEfectivos >= 1) Then
            .Value = "REZAGO EN EL DESARROLLO (AMARILLO)"
            .Interior.Color = RGB(255, 243, 205)
            .Font.Color = RGB(133, 100, 4)
        Else
            .Value = "DESARROLLO NORMAL (VERDE)"
            .Interior.Color = RGB(212, 237, 218)
            .Font.Color = RGB(21, 87, 36)
        End If
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ============================================================================
' 3. BOTÓN "3. GUARDAR EN HISTÓRICO" (32 COLUMNAS CON ÁREA DE EVALUACIÓN)
' ============================================================================
Public Sub GuardarEnHistorico()
    Dim dash As Worksheet, hist As Worksheet, hojaCalculo As Worksheet
    Dim resGlobal As String, dni As String, nombreNino As String, nombreMadre As String
    Dim tipoSeguro As String, direccionNino As String, referenciaNino As String, evaluadorEDI As String
    Dim celularNino As String, pobIndigena As String, perimetroCef As String, areaEvaluacion As String
    Dim strFAtencion As String, strFNac As String, grupoEDI As String, strGestHist As String
    Dim edadMeses As Double
    Dim ultimaFilaHist As Long, h As Long, i As Long
    Dim cod As String, tipoSeccion As String, subArea As String, resp As String, resBloque As String, itemTexto As String
    Dim idEval As String, nuevaFila As Long
    Const FILA_INI As Long = 24
    Const FILA_FIN As Long = 75
    
    Dim resMG As String, resMF As String, resLE As String, resSO As String, resCO As String
    Dim resAlarma As String, resNeuro As String, resRiesgo As String, resAlerta As String
    Dim detDevAct As String, detDevAnt As String, detAlmNeuro As String, detRiesgoAlerta As String
    
    Set dash = ThisWorkbook.Sheets("Dashboard")
    Set hist = ThisWorkbook.Sheets("Historico")
    Set hojaCalculo = ThisWorkbook.Sheets("Calculo")
    
    Dim headers As Variant
    headers = Array("ID_Evaluacion", "Fecha_Registro", "DNI", "Nombre_Nino", "Fecha_Nacimiento", _
                    "Edad_Gestacional", "Edad_Meses_Eval", "Fecha_Atencion", "Nombre_Madre", "Grupo_EDI", _
                    "Diagnostico_Global", "Res_Motriz_Gruesa", "Res_Motriz_Fina", "Res_Lenguaje", _
                    "Res_Social", "Res_Conocimiento", "Res_Senales_Alarma", "Res_Expl_Neurologica", _
                    "Res_Riesgo_Biologico", "Res_Senales_Alerta", "Detalle_Desarrollo_Actual", _
                    "Detalle_Grupo_Anterior", "Detalle_Alarma_Neuro", "Detalle_Riesgo_Alerta", _
                    "Tipo_Seguro", "Direccion", "Referencia", "Evaluador_EDI", _
                    "Celular", "Poblacion_Indigena", "Perimetro_Cefalico", "Area_Evaluacion")
    Dim colH As Integer
    For colH = 0 To UBound(headers)
        hist.Cells(1, colH + 1).Value = headers(colH)
    Next colH
    hist.Range("A1:AF1").Font.Bold = True
    
    resGlobal = Trim(CStr(dash.Range("C20").MergeArea.Cells(1, 1).Value))
    If resGlobal = "" Or InStr(resGlobal, "PENDIENTE") > 0 Or InStr(resGlobal, "NO APLICA") > 0 Then
        MsgBox "Debe tener un diagnóstico EDI válido calculado antes de guardar.", vbExclamation
        Exit Sub
    End If
    
    dni = Trim(CStr(dash.Range("B5").MergeArea.Cells(1, 1).Value))
    If dni = "" Then
        MsgBox "Por favor ingresa el DNI del niño en B5.", vbExclamation, "Campo Requerido"
        Exit Sub
    End If
    
    areaEvaluacion = UCase(Trim(CStr(dash.Range("F10").MergeArea.Cells(1, 1).Value)))
    If areaEvaluacion = "" Then
        MsgBox "Por favor selecciona el Área de Evaluación (INICIAL, SUBSECUENTE 1..5) en F10.", vbExclamation, "Campo Requerido"
        Exit Sub
    End If
    
    nombreNino = Trim(CStr(dash.Range("B6").MergeArea.Cells(1, 1).Value))
    nombreMadre = Trim(CStr(dash.Range("B7").MergeArea.Cells(1, 1).Value))
    celularNino = Trim(CStr(dash.Range("B8").MergeArea.Cells(1, 1).Value))
    evaluadorEDI = Trim(CStr(dash.Range("B9").MergeArea.Cells(1, 1).Value))
    tipoSeguro = Trim(CStr(dash.Range("B10").MergeArea.Cells(1, 1).Value))
    
    direccionNino = Trim(CStr(dash.Range("D9").MergeArea.Cells(1, 1).Value))
    referenciaNino = Trim(CStr(dash.Range("D10").MergeArea.Cells(1, 1).Value))
    
    strFAtencion = Format(CDate(dash.Range("F5").MergeArea.Cells(1, 1).Value), "yyyy-mm-dd")
    strFNac = Format(CDate(dash.Range("F6").MergeArea.Cells(1, 1).Value), "yyyy-mm-dd")
    pobIndigena = UCase(Trim(CStr(dash.Range("F8").MergeArea.Cells(1, 1).Value)))
    perimetroCef = Trim(CStr(dash.Range("F9").MergeArea.Cells(1, 1).Value))
    
    If Trim(CStr(dash.Range("F7").MergeArea.Cells(1, 1).Value)) = "" Or Val(dash.Range("F7").MergeArea.Cells(1, 1).Value) = 0 Then
        strGestHist = "A término"
    Else
        strGestHist = Trim(CStr(dash.Range("F7").MergeArea.Cells(1, 1).Value)) & " sem"
    End If
    
    grupoEDI = Trim(CStr(dash.Range("B15").MergeArea.Cells(1, 1).Value))
    edadMeses = Val(hojaCalculo.Range("F18").Value)
    
    ultimaFilaHist = hist.Cells(hist.Rows.Count, 1).End(xlUp).Row
    For h = 2 To ultimaFilaHist
        If Trim(CStr(hist.Cells(h, 3).Value)) = dni Then
            If Format(hist.Cells(h, 8).Value, "yyyy-mm-dd") = strFAtencion Then
                MsgBox "Ya existe un registro en el Histórico para el DNI (" & dni & ") en la Fecha de Atención (" & strFAtencion & ").", vbCritical, "Registro Duplicado"
                Exit Sub
            End If
            If UCase(Trim(CStr(hist.Cells(h, 32).Value))) = areaEvaluacion Then
                MsgBox "El niño con DNI (" & dni & ") ya tiene registrada la etapa '" & areaEvaluacion & "' en el Histórico.", vbCritical, "Etapa Duplicada"
                Exit Sub
            End If
        End If
    Next h
    
    resMG = "NO APLICA": resMF = "NO APLICA": resLE = "NO APLICA": resSO = "NO APLICA": resCO = "NO APLICA"
    resAlarma = "NO APLICA": resNeuro = "NO APLICA": resRiesgo = "NO APLICA": resAlerta = "NO APLICA"
    
    For i = FILA_INI To FILA_FIN
        cod = Trim(CStr(dash.Cells(i, 1).Value))
        If cod = "" Then GoTo SiguienteHist
        
        tipoSeccion = Trim(CStr(dash.Cells(i, 2).Value))
        subArea = Trim(CStr(dash.Cells(i, 3).Value))
        resp = Trim(CStr(dash.Cells(i, 6).Value))
        resBloque = Trim(CStr(dash.Cells(i, 8).MergeArea.Cells(1, 1).Value))
        
        If InStr(cod, "_ANT") = 0 Then
            If InStr(tipoSeccion, "Desarrollo") > 0 Or InStr(tipoSeccion, "Área") > 0 Then
                If InStr(LCase(subArea), "gruesa") > 0 Then resMG = resBloque
                If InStr(LCase(subArea), "fina") > 0 Then resMF = resBloque
                If InStr(LCase(subArea), "lenguaje") > 0 Then resLE = resBloque
                If InStr(LCase(subArea), "social") > 0 Then resSO = resBloque
                If InStr(LCase(subArea), "conocimiento") > 0 Then resCO = resBloque
            ElseIf InStr(tipoSeccion, "Alarma") > 0 Then
                resAlarma = resBloque
            ElseIf InStr(tipoSeccion, "Neurol") > 0 Then
                resNeuro = resBloque
            ElseIf InStr(tipoSeccion, "Riesgo") > 0 Then
                resRiesgo = resBloque
            ElseIf InStr(tipoSeccion, "Alerta") > 0 Then
                resAlerta = resBloque
            End If
        End If
        
        If resp <> "" Then
            itemTexto = cod & ": " & resp
            If InStr(cod, "_ANT") > 0 Then
                detDevAnt = IIf(detDevAnt = "", itemTexto, detDevAnt & " | " & itemTexto)
            ElseIf InStr(tipoSeccion, "Desarrollo") > 0 Or InStr(tipoSeccion, "Área") > 0 Then
                detDevAct = IIf(detDevAct = "", itemTexto, detDevAct & " | " & itemTexto)
            ElseIf InStr(tipoSeccion, "Alarma") > 0 Or InStr(tipoSeccion, "Neurol") > 0 Then
                detAlmNeuro = IIf(detAlmNeuro = "", itemTexto, detAlmNeuro & " | " & itemTexto)
            ElseIf InStr(tipoSeccion, "Riesgo") > 0 Or InStr(tipoSeccion, "Alerta") > 0 Then
                detRiesgoAlerta = IIf(detRiesgoAlerta = "", itemTexto, detRiesgoAlerta & " | " & itemTexto)
            End If
        End If
SiguienteHist:
    Next i
    
    nuevaFila = ultimaFilaHist + 1
    idEval = "EDI-" & Format(Now, "yyyymmddhhnnss")
    
    hist.Cells(nuevaFila, 1).Value = idEval
    hist.Cells(nuevaFila, 2).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    hist.Cells(nuevaFila, 3).Value = dni
    hist.Cells(nuevaFila, 4).Value = nombreNino
    hist.Cells(nuevaFila, 5).Value = strFNac
    hist.Cells(nuevaFila, 6).Value = strGestHist
    hist.Cells(nuevaFila, 7).Value = edadMeses
    hist.Cells(nuevaFila, 8).Value = strFAtencion
    hist.Cells(nuevaFila, 9).Value = nombreMadre
    hist.Cells(nuevaFila, 10).Value = grupoEDI
    hist.Cells(nuevaFila, 11).Value = resGlobal
    hist.Cells(nuevaFila, 12).Value = resMG
    hist.Cells(nuevaFila, 13).Value = resMF
    hist.Cells(nuevaFila, 14).Value = resLE
    hist.Cells(nuevaFila, 15).Value = resSO
    hist.Cells(nuevaFila, 16).Value = resCO
    hist.Cells(nuevaFila, 17).Value = resAlarma
    hist.Cells(nuevaFila, 18).Value = resNeuro
    hist.Cells(nuevaFila, 19).Value = resRiesgo
    hist.Cells(nuevaFila, 20).Value = resAlerta
    hist.Cells(nuevaFila, 21).Value = IIf(detDevAct = "", "Sin respuestas", detDevAct)
    hist.Cells(nuevaFila, 22).Value = IIf(detDevAnt = "", "No aplicado", detDevAnt)
    hist.Cells(nuevaFila, 23).Value = IIf(detAlmNeuro = "", "No aplicado", detAlmNeuro)
    hist.Cells(nuevaFila, 24).Value = IIf(detRiesgoAlerta = "", "No aplicado", detRiesgoAlerta)
    hist.Cells(nuevaFila, 25).Value = tipoSeguro
    hist.Cells(nuevaFila, 26).Value = direccionNino
    hist.Cells(nuevaFila, 27).Value = referenciaNino
    hist.Cells(nuevaFila, 28).Value = evaluadorEDI
    hist.Cells(nuevaFila, 29).Value = celularNino
    hist.Cells(nuevaFila, 30).Value = pobIndigena
    hist.Cells(nuevaFila, 31).Value = perimetroCef
    hist.Cells(nuevaFila, 32).Value = areaEvaluacion
    
    With hist.Range(hist.Cells(nuevaFila, 1), hist.Cells(nuevaFila, 32))
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    hist.Rows(nuevaFila).AutoFit
    
    MsgBox "¡Evaluación (" & areaEvaluacion & ") guardada en Histórico correctamente!" & vbCrLf & "ID: " & idEval, vbInformation, "Guardado Exitoso"
End Sub

' ============================================================================
' 4. BOTÓN OPCIONAL: LIMPIAR FORMULARIO
' ============================================================================
Public Sub LimpiarFormularioEDI()
    Dim dash As Worksheet, rCel As Range
    Const FILA_INI As Long = 24
    Const FILA_FIN As Long = 75
    Set dash = ThisWorkbook.Sheets("Dashboard")
    
    Application.EnableEvents = False
    For Each rCel In dash.Range("B5:B10,D9:D10,F5:F10,B15")
        rCel.MergeArea.ClearContents
    Next rCel
    
    With dash.Range("C20").MergeArea.Cells(1, 1)
        .Value = "PENDIENTE DE EVALUACIÓN"
        .Interior.Color = RGB(232, 236, 239)
        .Font.Color = RGB(0, 0, 0)
    End With
    
    For Each rCel In dash.Range("A" & FILA_INI & ":H" & FILA_FIN)
        If rCel.MergeCells Then rCel.MergeArea.UnMerge
    Next rCel
    dash.Range("A" & FILA_INI & ":H" & FILA_FIN).ClearContents
    dash.Range("A" & FILA_INI & ":H" & FILA_FIN).Interior.ColorIndex = xlNone
    dash.Range("A" & FILA_INI & ":H" & FILA_FIN).Borders.LineStyle = xlNone
    dash.Range("F" & FILA_INI & ":F" & FILA_FIN).Validation.Delete
    dash.Rows(FILA_INI & ":" & FILA_FIN).RowHeight = 20
    Application.EnableEvents = True
    
    MsgBox "Formulario limpiado. Listo para registrar un nuevo paciente.", vbInformation
End Sub

