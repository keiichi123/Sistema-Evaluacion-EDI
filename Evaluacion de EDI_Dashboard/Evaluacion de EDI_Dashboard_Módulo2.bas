Attribute VB_Name = "Módulo2"
Option Explicit

' ============================================================================
' GENERADOR FIEL DE LA FICHA OFICIAL EDI (SOLO VOLCADO DESDE HISTÓRICO)
' ============================================================================
Public Sub GenerarFichaOficialEDI()
    Dim dash As Worksheet, hist As Worksheet, ficha As Worksheet
    Dim dniDefecto As String, dniBuscado As String, dniFila As String
    Dim ultimaFilaHist As Long, r As Long, i As Long, j As Long
    Dim colEtapa As Integer, colBase As Integer
    Dim encontrados As Integer
    Dim filasPaciente() As Long
    Dim fechasPaciente() As Date
    Dim ordenEtapaPaciente() As Integer
    
    On Error Resume Next
    Set dash = ThisWorkbook.Sheets("Dashboard")
    Set hist = ThisWorkbook.Sheets("Historico")
    On Error GoTo 0
    
    If hist Is Nothing Then
        MsgBox "No se encontró la pestaña 'Historico'.", vbCritical
        Exit Sub
    End If
    
    dniDefecto = ""
    If Not dash Is Nothing Then
        dniDefecto = LimpiarTextoDNI(dash.Range("B5").MergeArea.Cells(1, 1).Value)
    End If
    
    dniBuscado = LimpiarTextoDNI(InputBox("Ingresa el DNI del niño para generar su Ficha Oficial EDI:", _
                                          "Generar Ficha Oficial EDI", dniDefecto))
    If dniBuscado = "" Then Exit Sub
    
    ' 1. Recorrer TODAS las filas con datos en 'Historico' para ese DNI
    ultimaFilaHist = hist.Cells(hist.Rows.Count, 3).End(xlUp).Row
    If hist.UsedRange.Rows.Count + hist.UsedRange.Row - 1 > ultimaFilaHist Then
        ultimaFilaHist = hist.UsedRange.Rows.Count + hist.UsedRange.Row - 1
    End If
    
    encontrados = 0
    For r = 2 To ultimaFilaHist
        dniFila = LimpiarTextoDNI(hist.Cells(r, 3).Value)
        If dniFila <> "" And dniFila = dniBuscado Then
            encontrados = encontrados + 1
            ReDim Preserve filasPaciente(1 To encontrados)
            ReDim Preserve fechasPaciente(1 To encontrados)
            ReDim Preserve ordenEtapaPaciente(1 To encontrados)
            
            filasPaciente(encontrados) = r
            
            If IsDate(hist.Cells(r, 8).Value) Then
                fechasPaciente(encontrados) = CDate(hist.Cells(r, 8).Value)
            Else
                fechasPaciente(encontrados) = DateSerial(1900, 1, encontrados)
            End If
            
            ordenEtapaPaciente(encontrados) = ObtenerIndiceEtapa(CStr(hist.Cells(r, 32).Value))
        End If
    Next r
    
    If encontrados = 0 Then
        MsgBox "No se encontraron evaluaciones guardadas en 'Historico' para el DNI: " & dniBuscado, vbExclamation, "Sin Registros"
        Exit Sub
    End If
    
    ' 2. Ordenar cronológicamente y por etapa (Inicial -> Subsecuente 1 -> ... -> Subsecuente 5)
    Dim tmpFila As Long, tmpFecha As Date, tmpOrd As Integer
    Dim debeIntercambiar As Boolean
    For i = 1 To encontrados - 1
        For j = i + 1 To encontrados
            debeIntercambiar = False
            If fechasPaciente(j) < fechasPaciente(i) Then
                debeIntercambiar = True
            ElseIf fechasPaciente(j) = fechasPaciente(i) Then
                If ordenEtapaPaciente(j) >= 0 And ordenEtapaPaciente(i) >= 0 And ordenEtapaPaciente(j) <> ordenEtapaPaciente(i) Then
                    If ordenEtapaPaciente(j) < ordenEtapaPaciente(i) Then debeIntercambiar = True
                ElseIf filasPaciente(j) < filasPaciente(i) Then
                    debeIntercambiar = True
                End If
            End If
            
            If debeIntercambiar Then
                tmpFecha = fechasPaciente(i): fechasPaciente(i) = fechasPaciente(j): fechasPaciente(j) = tmpFecha
                tmpFila = filasPaciente(i): filasPaciente(i) = filasPaciente(j): filasPaciente(j) = tmpFila
                tmpOrd = ordenEtapaPaciente(i): ordenEtapaPaciente(i) = ordenEtapaPaciente(j): ordenEtapaPaciente(j) = tmpOrd
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    ThisWorkbook.Sheets("Ficha_EDI").Delete
    On Error GoTo 0
    
    Set ficha = ThisWorkbook.Sheets.Add(After:=hist)
    ficha.Name = "Ficha_EDI"
    
    ' 3. Construir la plantilla oficial exacta
    ConstruirPlantillaFichaOficial ficha
    
    ' 4. Volcar TODAS las evaluaciones históricas del niño respetando la etapa exacta de cada registro
    Dim cabeceraLlena As Boolean
    Dim etapaOcupada(0 To 5) As Boolean
    Dim etapaEval As Integer
    cabeceraLlena = False
    
    For i = 1 To encontrados
        r = filasPaciente(i)
        etapaEval = ordenEtapaPaciente(i)
        
        If etapaEval >= 0 And etapaEval <= 5 And Not etapaOcupada(etapaEval) Then
            colEtapa = etapaEval
        Else
            colEtapa = ObtenerSiguienteEtapaLibre(etapaOcupada)
        End If
        
        If colEtapa >= 0 And colEtapa <= 5 Then
            etapaOcupada(colEtapa) = True
            
            If Not cabeceraLlena Or colEtapa = 0 Then
                ficha.Range("H3").Value = CStr(hist.Cells(r, 4).Value)
                ficha.Range("AE3").Value = CStr(hist.Cells(r, 25).Value)
                ficha.Range("H4").Value = CStr(hist.Cells(r, 9).Value)
                ficha.Range("H5").Value = CStr(hist.Cells(r, 26).Value)
                ficha.Range("AE5").Value = CStr(hist.Cells(r, 29).Value)
                ficha.Range("H6").Value = CStr(hist.Cells(r, 27).Value)
                
                ficha.Range("AI6").Value = ""
                ficha.Range("AL6").Value = ""
                If UCase(Trim(CStr(hist.Cells(r, 30).Value))) = "SI" Then
                    ficha.Range("AI6").Value = "X"
                ElseIf UCase(Trim(CStr(hist.Cells(r, 30).Value))) = "NO" Then
                    ficha.Range("AL6").Value = "X"
                End If
                
                If IsDate(hist.Cells(r, 5).Value) Then
                    Dim fN As Date
                    fN = CDate(hist.Cells(r, 5).Value)
                    ficha.Range("A9").Value = Format(Day(fN), "00")
                    ficha.Range("B9").Value = Format(Month(fN), "00")
                    ficha.Range("C9").Value = Right(CStr(Year(fN)), 2)
                End If
                
                ficha.Range("A12").Value = CStr(hist.Cells(r, 6).Value)
                cabeceraLlena = True
            End If
            
            If Trim(CStr(hist.Cells(r, 28).Value)) <> "" Then
                ficha.Range("A14").Value = CStr(hist.Cells(r, 28).Value)
            End If
            
            colBase = 4 + (colEtapa * 6)
            
            If IsDate(hist.Cells(r, 8).Value) Then
                ficha.Cells(8, colBase + 2).Value = Format(CDate(hist.Cells(r, 8).Value), "dd/mm/yy")
            End If
            
            Dim grpNum As Integer
            grpNum = ExtraerGrupoFicha(CStr(hist.Cells(r, 10).Value))
            If grpNum > 0 Then ficha.Cells(12, colBase + 4).Value = grpNum
            
            If IsDate(hist.Cells(r, 5).Value) And IsDate(hist.Cells(r, 8).Value) Then
                Dim fNacD As Date, fAteD As Date, semG As Integer
                fNacD = CDate(hist.Cells(r, 5).Value)
                fAteD = CDate(hist.Cells(r, 8).Value)
                semG = Val(hist.Cells(r, 6).Value)
                
                Dim aC As Integer, mC As Integer, dC As Integer
                Dim aR As Integer, mR As Integer, dR As Integer
                CalcularDesgloseEdades fNacD, fAteD, semG, aC, mC, dC, aR, mR, dR
                
                ficha.Cells(9, colBase + 2).Value = aC & "a " & mC & "m " & dC & "d"
                If semG > 0 And semG < 37 And aC < 2 Then
                    ficha.Cells(10, colBase + 2).Value = aR & "a " & mR & "m " & dR & "d"
                Else
                    ficha.Cells(10, colBase + 2).Value = "No aplica"
                End If
            End If
            
            ' Áreas del Desarrollo (MG=18, MF=22, LE=26, SO=30, CO=34)
            Dim detAct As String, detAnt As String
            detAct = CStr(hist.Cells(r, 21).Value)
            detAnt = CStr(hist.Cells(r, 22).Value)
            
            MarcarAreaEnFicha ficha, colBase, 18, "MG", detAct, detAnt, CStr(hist.Cells(r, 12).Value), grpNum
            MarcarAreaEnFicha ficha, colBase, 22, "MF", detAct, detAnt, CStr(hist.Cells(r, 13).Value), grpNum
            MarcarAreaEnFicha ficha, colBase, 26, "LEN", detAct, detAnt, CStr(hist.Cells(r, 14).Value), grpNum
            MarcarAreaEnFicha ficha, colBase, 30, "SOC", detAct, detAnt, CStr(hist.Cells(r, 15).Value), grpNum
            MarcarAreaEnFicha ficha, colBase, 34, "CON", detAct, detAnt, CStr(hist.Cells(r, 16).Value), grpNum
            
            ' Señales de Alarma (Filas 40..43)
            Dim detAlmNeu As String
            detAlmNeu = CStr(hist.Cells(r, 23).Value)
            MarcarAlarmaNeuroEnFicha ficha, colBase, 40, "ALM", 4, detAlmNeu, CStr(hist.Cells(r, 17).Value), ""
            
            ' Exploración Neurológica (Filas 46..48) + Perímetro Cefálico
            Dim pcVal As String, textoPC As String
            pcVal = Trim(CStr(hist.Cells(r, 31).Value))
            textoPC = IIf(pcVal <> "", "PC " & pcVal & " cm", "PC --- cm")
            MarcarAlarmaNeuroEnFicha ficha, colBase, 46, "EN", 3, detAlmNeu, CStr(hist.Cells(r, 18).Value), textoPC
            
            ' Resultado Global (Fila 50: Verde -> Amarillo -> Rojo, igual que la ficha física)
            Dim diagGlob As String
            diagGlob = UCase(CStr(hist.Cells(r, 11).Value))
            If InStr(diagGlob, "VERDE") > 0 Then
                ficha.Cells(50, colBase + 1).Value = "X"
            ElseIf InStr(diagGlob, "AMARILLO") > 0 Then
                ficha.Cells(50, colBase + 2).Value = "X"
            ElseIf InStr(diagGlob, "ROJO") > 0 Then
                ficha.Cells(50, colBase + 3).Value = "X"
            End If
            
            ' FRB (solo en Inicial) y Señales de Alerta (cuenta amarillos -> "1 A", "2 A", o blanco si es 0)
            Dim detRiesgoAle As String
            detRiesgoAle = CStr(hist.Cells(r, 24).Value)
            If colEtapa = 0 Then
                MarcarFRBEnFicha ficha, detRiesgoAle
            End If
            MarcarAlertaEnFicha ficha, colBase, detRiesgoAle
        End If
    Next i
    
    ficha.Activate
    ficha.Range("A1").Select
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "¡Se cargaron " & encontrados & " evaluación(es) histórica(s) en la Ficha Oficial EDI para " & ficha.Range("H3").Value & " (DNI: " & dniBuscado & ")!", vbInformation, "Ficha Oficial Completa"
End Sub

' ============================================================================
' CONSTRUYE LA PLANTILLA OFICIAL (TÍTULOS CELESTE PASTEL Y AMARILLO PASTEL)
' ============================================================================
Private Sub ConstruirPlantillaFichaOficial(ByVal ws As Worksheet)
    Dim c As Integer, r As Integer, et As Integer, cb As Integer
    Dim etapas As Variant
    Dim colRojoOficial As Long, colAmarilloOficial As Long, colVerdeOficial As Long
    Dim colCelestePastel As Long, colAmarilloPastel As Long
    
    colRojoOficial = RGB(217, 83, 79)
    colAmarilloOficial = RGB(244, 240, 66)
    colVerdeOficial = RGB(84, 178, 108)
    colCelestePastel = RGB(217, 226, 243)
    colAmarilloPastel = RGB(242, 242, 204)
    
    etapas = Array("Inicial", "Subsecuente 1", "Subsecuente 2", "Subsecuente 3", "Subsecuente 4", "Subsecuente 5")
    
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 8
    ws.Cells.Font.Color = RGB(0, 0, 0)
    
    ws.Columns("A:C").ColumnWidth = 3.6
    For et = 0 To 5
        cb = 4 + (et * 6)
        ws.Columns(cb).ColumnWidth = 2.8
        ws.Columns(cb + 1).ColumnWidth = 3.2
        ws.Columns(cb + 2).ColumnWidth = 3.2
        ws.Columns(cb + 3).ColumnWidth = 2.8
        ws.Columns(cb + 4).ColumnWidth = 3.2
        ws.Columns(cb + 5).ColumnWidth = 3.2
    Next et
    
    ws.Range("A1:N1").Merge: ws.Range("A1").Value = "FORMATO DE APLICACIÓN"
    ws.Range("O1:AM1").Merge: ws.Range("O1").Value = "PRUEBA DE EVALUACIÓN DEL DESARROLLO INFANTIL (EDI)"
    With ws.Range("A1:AM1")
        .Font.Bold = True: .Font.Size = 10
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = colCelestePastel
    End With
    
    ws.Range("A3:G3").Merge: ws.Range("A3").Value = "Nombre de la niña o niño"
    ws.Range("H3:Z3").Merge
    ws.Range("AA3:AD3").Merge: ws.Range("AA3").Value = "Tipo de seguro"
    ws.Range("AE3:AM3").Merge
    
    ws.Range("A4:G4").Merge: ws.Range("A4").Value = "Nombre de la madre/padre/cuidador"
    ws.Range("H4:AM4").Merge
    
    ws.Range("A5:G5").Merge: ws.Range("A5").Value = "Dirección del niño"
    ws.Range("H5:Z5").Merge
    ws.Range("AA5:AD5").Merge: ws.Range("AA5").Value = "Nº Celular"
    ws.Range("AE5:AM5").Merge
    
    ws.Range("A6:G6").Merge: ws.Range("A6").Value = "REFERENCIA"
    ws.Range("H6:Z6").Merge
    ws.Range("AA6:AF6").Merge: ws.Range("AA6").Value = "Población indígena"
    ws.Range("AG6:AH6").Merge: ws.Range("AG6").Value = "SI"
    ws.Range("AJ6:AK6").Merge: ws.Range("AJ6").Value = "NO"
    ws.Range("AG6:AM6").HorizontalAlignment = xlCenter
    
    ws.Range("A3:G6,AA3:AD3,AA5:AD5,AA6:AF6").Font.Bold = True
    ws.Range("H3:Z6,AE3:AM5,AI6,AL6").Font.Bold = True
    
    ws.Range("A7:AM7").Merge: ws.Range("A7").Value = "SEGUIMIENTO AL DESARROLLO"
    With ws.Range("A7:AM7")
        .Font.Bold = True: .HorizontalAlignment = xlCenter: .Interior.Color = colCelestePastel
    End With
    
    ws.Range("A8:C8").Merge: ws.Range("A8").Value = "Fecha Nac."
    ws.Range("A10").Value = "dd": ws.Range("B10").Value = "mm": ws.Range("C10").Value = "aa"
    ws.Range("A11:C11").Merge: ws.Range("A11").Value = "Sem. De Gestación"
    ws.Range("A12:C12").Merge
    ws.Range("A8:C12").HorizontalAlignment = xlCenter: ws.Range("A8:C12").Font.Bold = True
    
    For et = 0 To 5
        cb = 4 + (et * 6)
        ws.Range(ws.Cells(8, cb), ws.Cells(8, cb + 5)).Merge
        ws.Cells(8, cb).Value = etapas(et)
        ws.Cells(8, cb).Font.Bold = True: ws.Cells(8, cb).HorizontalAlignment = xlCenter
        ws.Cells(8, cb).Interior.Color = RGB(255, 255, 255)
        
        ws.Range(ws.Cells(9, cb), ws.Cells(9, cb + 1)).Merge: ws.Cells(9, cb).Value = "Edad"
        ws.Range(ws.Cells(9, cb + 2), ws.Cells(9, cb + 5)).Merge
        
        ws.Range(ws.Cells(10, cb), ws.Cells(10, cb + 1)).Merge: ws.Cells(10, cb).Value = "E.Corr"
        ws.Range(ws.Cells(10, cb + 2), ws.Cells(10, cb + 5)).Merge
        
        ws.Range(ws.Cells(11, cb), ws.Cells(11, cb + 5)).Merge: ws.Cells(11, cb).Value = "F. Atención:"
        ws.Range(ws.Cells(12, cb), ws.Cells(12, cb + 3)).Merge: ws.Cells(12, cb).Value = "Grupo de edad"
        ws.Range(ws.Cells(12, cb + 4), ws.Cells(12, cb + 5)).Merge
        ws.Range(ws.Cells(12, cb + 4), ws.Cells(12, cb + 5)).Font.Bold = True
        ws.Range(ws.Cells(12, cb + 4), ws.Cells(12, cb + 5)).HorizontalAlignment = xlCenter
    Next et
    
    ws.Range("A13:AM13").Merge: ws.Range("A13").Value = "NOMBRE COMPLETO DE LA PERSONA QUE APLICA LA PRUEBA EDI"
    ws.Range("A13:AM13").Font.Bold = True: ws.Range("A13:AM13").HorizontalAlignment = xlCenter
    ws.Range("A13:AM13").Interior.Color = colCelestePastel
    ws.Range("A14:AM14").Merge: ws.Range("A14").Font.Bold = True
    
    ws.Range("A15:AM15").Merge: ws.Range("A15").Value = "ÁREAS DEL DESARROLLO"
    ws.Range("A15:AM15").Font.Bold = True: ws.Range("A15:AM15").HorizontalAlignment = xlCenter
    ws.Range("A15:AM15").Interior.Color = colAmarilloPastel
    
    ws.Range("A16:C17").Merge
    For et = 0 To 5
        cb = 4 + (et * 6)
        ws.Range(ws.Cells(16, cb), ws.Cells(16, cb + 5)).Merge
        ws.Cells(16, cb).Value = etapas(et)
        ws.Cells(16, cb).Font.Bold = True: ws.Cells(16, cb).HorizontalAlignment = xlCenter
        
        ws.Range(ws.Cells(17, cb), ws.Cells(17, cb + 3)).Merge
        ws.Range(ws.Cells(17, cb + 4), ws.Cells(17, cb + 5)).Merge
        ws.Cells(17, cb + 4).Value = "Grupo anterior"
        ws.Cells(17, cb + 4).Font.Bold = True: ws.Cells(17, cb + 4).Font.Size = 6.8
        ws.Cells(17, cb + 4).HorizontalAlignment = xlCenter: ws.Cells(17, cb + 4).WrapText = True
    Next et
    
    Dim areasCod As Variant, aIdx As Integer, fBase As Integer
    areasCod = Array("MG", "MF", "LE", "SO", "CO")
    
    For aIdx = 0 To 4
        fBase = 18 + (aIdx * 4)
        ws.Range(ws.Cells(fBase, 1), ws.Cells(fBase + 3, 3)).Merge
        ws.Cells(fBase, 1).Value = areasCod(aIdx)
        ws.Cells(fBase, 1).Font.Bold = True: ws.Cells(fBase, 1).Font.Size = 10
        ws.Cells(fBase, 1).HorizontalAlignment = xlCenter: ws.Cells(fBase, 1).VerticalAlignment = xlCenter
        
        For et = 0 To 5
            cb = 4 + (et * 6)
            ws.Cells(fBase, cb).Value = "1"
            ws.Cells(fBase + 1, cb).Value = "2"
            ws.Cells(fBase + 2, cb).Value = "3"
            
            For r = 0 To 2
                ws.Cells(fBase + r, cb + 1).Interior.Color = colAmarilloOficial
                ws.Cells(fBase + r, cb + 2).Interior.Color = colVerdeOficial
                ws.Cells(fBase + r, cb + 4).Interior.Color = colRojoOficial
                ws.Cells(fBase + r, cb + 5).Interior.Color = colAmarilloOficial
            Next r
            
            ws.Range(ws.Cells(fBase, cb + 3), ws.Cells(fBase + 2, cb + 3)).Merge
            ws.Range(ws.Cells(fBase + 3, cb), ws.Cells(fBase + 3, cb + 2)).Merge
            ws.Cells(fBase + 3, cb).Value = "Resultado"
            ws.Range(ws.Cells(fBase + 3, cb + 3), ws.Cells(fBase + 3, cb + 5)).Merge
            
            ws.Range(ws.Cells(fBase, cb), ws.Cells(fBase + 3, cb + 5)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(fBase + 3, cb), ws.Cells(fBase + 3, cb + 5)).Font.Bold = True
        Next et
    Next aIdx
    
    ws.Range("A38:AM38").Merge: ws.Range("A38").Value = "SEÑALES DE ALARMA"
    ws.Range("A38:AM38").Font.Bold = True: ws.Range("A38:AM38").HorizontalAlignment = xlCenter
    ws.Range("A38:AM38").Interior.Color = colAmarilloPastel
    
    ws.Range("A39:C39").Merge
    For et = 0 To 5
        cb = 4 + (et * 6)
        ws.Range(ws.Cells(39, cb), ws.Cells(39, cb + 5)).Merge
        ws.Cells(39, cb).Value = etapas(et)
        ws.Cells(39, cb).Font.Bold = True: ws.Cells(39, cb).HorizontalAlignment = xlCenter
    Next et
    
    ws.Range("A40:C43").Merge: ws.Range("A40").Value = "ALM"
    ws.Range("A40:C43").Font.Bold = True: ws.Range("A40:C43").HorizontalAlignment = xlCenter
    
    For et = 0 To 5
        cb = 4 + (et * 6)
        For r = 0 To 3
            ws.Cells(40 + r, cb).Value = r + 1
            ws.Cells(40 + r, cb + 1).Interior.Color = colRojoOficial
            ws.Cells(40 + r, cb + 2).Interior.Color = colVerdeOficial
        Next r
        ws.Range(ws.Cells(40, cb + 3), ws.Cells(43, cb + 3)).Merge
        ws.Cells(40, cb + 3).Value = "Resultado"
        ws.Cells(40, cb + 3).Orientation = 90
        ws.Range(ws.Cells(40, cb + 4), ws.Cells(43, cb + 5)).Merge
        ws.Range(ws.Cells(40, cb), ws.Cells(43, cb + 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(40, cb + 3), ws.Cells(43, cb + 5)).Font.Bold = True
    Next et
    
    ws.Range("A44:AM44").Merge: ws.Range("A44").Value = "EXPLORACIÓN NEUROLÓGICA"
    ws.Range("A44:AM44").Font.Bold = True: ws.Range("A44:AM44").HorizontalAlignment = xlCenter
    ws.Range("A44:AM44").Interior.Color = colAmarilloPastel
    
    ws.Range("A45:C45").Merge
    For et = 0 To 5
        cb = 4 + (et * 6)
        ws.Range(ws.Cells(45, cb), ws.Cells(45, cb + 5)).Merge
        ws.Cells(45, cb).Value = etapas(et)
        ws.Cells(45, cb).Font.Bold = True: ws.Cells(45, cb).HorizontalAlignment = xlCenter
    Next et
    
    ws.Range("A46:C48").Merge: ws.Range("A46").Value = "EN"
    ws.Range("A46:C48").Font.Bold = True: ws.Range("A46:C48").HorizontalAlignment = xlCenter
    
    For et = 0 To 5
        cb = 4 + (et * 6)
        For r = 0 To 2
            ws.Cells(46 + r, cb).Value = r + 1
            ws.Cells(46 + r, cb + 1).Interior.Color = colRojoOficial
            ws.Cells(46 + r, cb + 2).Interior.Color = colVerdeOficial
        Next r
        ws.Range(ws.Cells(46, cb + 3), ws.Cells(48, cb + 3)).Merge
        ws.Cells(46, cb + 3).Value = "PC --- cm"
        ws.Cells(46, cb + 3).Orientation = 90
        ws.Range(ws.Cells(46, cb + 4), ws.Cells(48, cb + 4)).Merge
        ws.Cells(46, cb + 4).Value = "Resultado"
        ws.Cells(46, cb + 4).Orientation = 90
        ws.Range(ws.Cells(46, cb + 5), ws.Cells(48, cb + 5)).Merge
        ws.Range(ws.Cells(46, cb), ws.Cells(48, cb + 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(46, cb + 3), ws.Cells(48, cb + 5)).Font.Bold = True
    Next et
    
    ws.Range("A49:AM49").Merge: ws.Range("A49").Value = "RESULTADO GLOBAL DE LA EVALUACIÓN"
    ws.Range("A49:AM49").Font.Bold = True: ws.Range("A49:AM49").HorizontalAlignment = xlCenter
    ws.Range("A49:AM49").Interior.Color = colCelestePastel
    
    ws.Range("A50:C50").Merge: ws.Range("A50").Value = "GLOBAL"
    ws.Range("A50:C50").Font.Bold = True: ws.Range("A50:C50").HorizontalAlignment = xlCenter
    
    For et = 0 To 5
        cb = 4 + (et * 6)
        ' Orden físico oficial: Verde -> Amarillo -> Rojo
        ws.Cells(50, cb + 1).Interior.Color = colVerdeOficial
        ws.Cells(50, cb + 2).Interior.Color = colAmarilloOficial
        ws.Cells(50, cb + 3).Interior.Color = colRojoOficial
        ws.Range(ws.Cells(50, cb + 4), ws.Cells(50, cb + 5)).Merge
        ws.Range(ws.Cells(50, cb), ws.Cells(50, cb + 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(50, cb), ws.Cells(50, cb + 5)).Font.Bold = True
    Next et
    
    ws.Range("A51:AM51").Merge
    ws.Range("A51").Value = "FACTORES DE RIESGO BIOLÓGICOS Y SEÑALES DE ALERTA SON OPCIONALES A PARTIR DEL GRUPO 5 (5 meses) YA QUE NO MODIFICA LA CALIFICACIÓN GLOBAL"
    ws.Range("A51:AM51").Font.Bold = True: ws.Range("A51:AM51").Font.Size = 7
    ws.Range("A51:AM51").HorizontalAlignment = xlCenter
    
    ws.Range("A52:C52").Merge: ws.Range("A52").Value = "FRB"
    ws.Range("A52:C52").Font.Bold = True: ws.Range("A52:C52").HorizontalAlignment = xlCenter
    ws.Range("A52:C52").Interior.Color = colCelestePastel
    ws.Range("D52:AM52").Merge: ws.Range("D52").Value = "SEÑALES DE ALERTA"
    ws.Range("D52:AM52").Font.Bold = True: ws.Range("D52:AM52").HorizontalAlignment = xlCenter
    ws.Range("D52:AM52").Interior.Color = colCelestePastel
    
    For r = 1 To 7
        ws.Cells(52 + r, 1).Value = r
        ws.Cells(52 + r, 1).HorizontalAlignment = xlCenter
        ws.Cells(52 + r, 2).Interior.Color = colAmarilloOficial
        ws.Cells(52 + r, 3).Interior.Color = colVerdeOficial
    Next r
    ws.Cells(60, 1).Value = "Resultado": ws.Cells(60, 1).Font.Bold = True: ws.Cells(60, 1).Font.Size = 6
    ws.Range("B60:C60").Merge: ws.Range("B60:C60").Font.Bold = True: ws.Range("B60:C60").HorizontalAlignment = xlCenter
    
    For et = 0 To 5
        cb = 4 + (et * 6)
        For r = 1 To 8
            ws.Cells(52 + r, cb).Value = r
            ws.Cells(52 + r, cb).HorizontalAlignment = xlCenter
            ws.Cells(52 + r, cb + 1).Interior.Color = colAmarilloOficial
            ws.Cells(52 + r, cb + 2).Interior.Color = colVerdeOficial
        Next r
        ws.Range(ws.Cells(53, cb + 3), ws.Cells(60, cb + 3)).Merge
        ws.Cells(53, cb + 3).Value = "Resultado"
        ws.Cells(53, cb + 3).Orientation = 90
        ws.Range(ws.Cells(53, cb + 4), ws.Cells(60, cb + 5)).Merge
        ws.Range(ws.Cells(53, cb), ws.Cells(60, cb + 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(53, cb + 3), ws.Cells(60, cb + 5)).Font.Bold = True
    Next et
    
    With ws.Range("A1:AM60")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
    End With
    
    ws.Rows("1:60").RowHeight = 14.5
    ws.Rows("1").RowHeight = 18
    ws.Rows("17").RowHeight = 20
    
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End Sub

' ============================================================================
' FUNCIONES AUXILIARES (SOLO LECTURA Y VOLCADO)
' ============================================================================
Private Function LimpiarTextoDNI(ByVal valCelda As Variant) As String
    Dim s As String
    s = Trim(CStr(valCelda))
    If Right(s, 2) = ".0" Then s = Left(s, Len(s) - 2)
    LimpiarTextoDNI = s
End Function

Private Function ObtenerIndiceEtapa(ByVal etapaTexto As String) As Integer
    Dim u As String
    u = UCase(Trim(etapaTexto))
    If InStr(u, "INICIAL") > 0 Then ObtenerIndiceEtapa = 0: Exit Function
    If InStr(u, "1") > 0 Then ObtenerIndiceEtapa = 1: Exit Function
    If InStr(u, "2") > 0 Then ObtenerIndiceEtapa = 2: Exit Function
    If InStr(u, "3") > 0 Then ObtenerIndiceEtapa = 3: Exit Function
    If InStr(u, "4") > 0 Then ObtenerIndiceEtapa = 4: Exit Function
    If InStr(u, "5") > 0 Then ObtenerIndiceEtapa = 5: Exit Function
    ObtenerIndiceEtapa = -1
End Function

Private Function ObtenerSiguienteEtapaLibre(ByRef etapaOcupada() As Boolean) As Integer
    Dim k As Integer
    For k = 0 To 5
        If Not etapaOcupada(k) Then
            ObtenerSiguienteEtapaLibre = k
            Exit Function
        End If
    Next k
    ObtenerSiguienteEtapaLibre = -1
End Function

Private Function ExtraerGrupoFicha(ByVal texto As String) As Integer
    Dim i As Integer, ch As String, dig As String
    For i = 1 To Len(texto)
        ch = Mid(texto, i, 1)
        If ch >= "0" And ch <= "9" Then
            dig = dig & ch
        ElseIf dig <> "" Then
            Exit For
        End If
    Next i
    ExtraerGrupoFicha = IIf(dig = "", 0, CInt(dig))
End Function

Private Function BuscarRespItem(ByVal cadenaDetalle As String, ByVal patronCod As String) As String
    Dim partes() As String, p As Variant, kv() As String, codLimpio As String
    BuscarRespItem = ""
    If cadenaDetalle = "" Or cadenaDetalle = "No aplicado" Or cadenaDetalle = "Sin respuestas" Then Exit Function
    partes = Split(cadenaDetalle, "|")
    For Each p In partes
        kv = Split(Trim(CStr(p)), ":")
        If UBound(kv) >= 1 Then
            codLimpio = UCase(Trim(kv(0)))
            If InStr(codLimpio, UCase(patronCod)) > 0 Then
                BuscarRespItem = UCase(Trim(kv(1)))
                Exit Function
            End If
        End If
    Next p
End Function

Private Function InicialResultado(ByVal resTexto As String) As String
    Dim u As String
    u = UCase(Trim(resTexto))
    If InStr(u, "VERDE") > 0 Then InicialResultado = "V": Exit Function
    If InStr(u, "AMARILLO") > 0 Then InicialResultado = "A": Exit Function
    If InStr(u, "ROJO") > 0 Then InicialResultado = "R": Exit Function
    InicialResultado = "N/A"
End Function

Private Sub PintarCeldaResultadoColor(ByVal rng As Range, ByVal letraRes As String)
    rng.Value = letraRes
    Select Case letraRes
        Case "V": rng.Interior.Color = RGB(84, 178, 108)
        Case "A": rng.Interior.Color = RGB(244, 240, 66)
        Case "R": rng.Interior.Color = RGB(217, 83, 79)
    End Select
End Sub

Private Sub MarcarAreaEnFicha(ByVal ws As Worksheet, ByVal cb As Integer, ByVal fBase As Integer, _
                              ByVal codArea As String, ByVal detAct As String, ByVal detAnt As String, _
                              ByVal resArea As String, ByVal grpNum As Integer)
    Dim itemIdx As Integer, respA As String, respAnt As String
    
    If grpNum = 1 Then
        For itemIdx = 1 To 3
            ws.Cells(fBase + itemIdx - 1, cb + 1).Interior.Color = RGB(217, 83, 79)
        Next itemIdx
    End If
    
    For itemIdx = 1 To 3
        respA = BuscarRespItem(detAct, "_" & codArea & "_0" & itemIdx)
        If respA = "" Then respA = BuscarRespItem(detAct, "_" & codArea & "_" & itemIdx)
        
        If respA = "SI" Then
            ws.Cells(fBase + itemIdx - 1, cb + 2).Value = "X"
        ElseIf respA = "NO" Then
            ws.Cells(fBase + itemIdx - 1, cb + 1).Value = "X"
        End If
        
        respAnt = BuscarRespItem(detAnt, "_" & codArea & "_0" & itemIdx & "_ANT")
        If respAnt = "" Then respAnt = BuscarRespItem(detAnt, "_" & codArea & "_" & itemIdx & "_ANT")
        If respAnt = "SI" Then
            ws.Cells(fBase + itemIdx - 1, cb + 5).Value = "X"
        ElseIf respAnt = "NO" Then
            ws.Cells(fBase + itemIdx - 1, cb + 4).Value = "X"
        End If
    Next itemIdx
    
    PintarCeldaResultadoColor ws.Cells(fBase + 3, cb + 3), InicialResultado(resArea)
End Sub

Private Sub MarcarAlarmaNeuroEnFicha(ByVal ws As Worksheet, ByVal cb As Integer, ByVal fBase As Integer, _
                                     ByVal codSec As String, ByVal totalItems As Integer, _
                                     ByVal detAlmNeu As String, ByVal resSec As String, _
                                     ByVal textoLateral As String)
    Dim itemIdx As Integer, resp As String, letraRes As String
    For itemIdx = 1 To totalItems
        resp = BuscarRespItem(detAlmNeu, "_" & codSec & "_0" & itemIdx)
        If resp = "" Then resp = BuscarRespItem(detAlmNeu, "_" & codSec & "_" & itemIdx)
        If resp = "SI" Then
            ws.Cells(fBase + itemIdx - 1, cb + 1).Value = "X"
        ElseIf resp = "NO" Then
            ws.Cells(fBase + itemIdx - 1, cb + 2).Value = "X"
        End If
    Next itemIdx
    
    letraRes = InicialResultado(resSec)
    If codSec = "ALM" Then
        PintarCeldaResultadoColor ws.Cells(fBase, cb + 4), letraRes
    Else
        ws.Cells(fBase, cb + 3).Value = textoLateral
        PintarCeldaResultadoColor ws.Cells(fBase, cb + 5), letraRes
    End If
End Sub

Private Sub MarcarFRBEnFicha(ByVal ws As Worksheet, ByVal detRiesgoAle As String)
    Dim itemIdx As Integer, resp As String, amarillos As Integer
    amarillos = 0
    For itemIdx = 1 To 7
        resp = BuscarRespItem(detRiesgoAle, "_FRB_0" & itemIdx)
        If resp = "" Then resp = BuscarRespItem(detRiesgoAle, "_FRB_" & itemIdx)
        If resp = "SI" Then
            ws.Cells(52 + itemIdx, 2).Value = "X"
            amarillos = amarillos + 1
        ElseIf resp = "NO" Then
            ws.Cells(52 + itemIdx, 3).Value = "X"
        End If
    Next itemIdx
    ws.Cells(60, 2).NumberFormat = "@"
    If amarillos > 0 Then
        ws.Cells(60, 2).Value = amarillos & " A"
    Else
        ws.Cells(60, 2).Value = ""
    End If
End Sub

Private Sub MarcarAlertaEnFicha(ByVal ws As Worksheet, ByVal cb As Integer, ByVal detRiesgoAle As String)
    Dim itemIdx As Integer, resp As String, amarillos As Integer
    amarillos = 0
    For itemIdx = 1 To 8
        resp = BuscarRespItem(detRiesgoAle, "_ALE_0" & itemIdx)
        If resp = "" Then resp = BuscarRespItem(detRiesgoAle, "_ALE_" & itemIdx)
        If resp = "SI" Then
            ws.Cells(52 + itemIdx, cb + 1).Value = "X"
            amarillos = amarillos + 1
        ElseIf resp = "NO" Then
            ws.Cells(52 + itemIdx, cb + 2).Value = "X"
        End If
    Next itemIdx
    ws.Cells(53, cb + 4).NumberFormat = "@"
    If amarillos > 0 Then
        ws.Cells(53, cb + 4).Value = amarillos & " A"
    Else
        ws.Cells(53, cb + 4).Value = ""
    End If
End Sub

Private Sub CalcularDesgloseEdades(ByVal fNac As Date, ByVal fAte As Date, ByVal semGest As Integer, _
                                   ByRef aC As Integer, ByRef mC As Integer, ByRef dC As Integer, _
                                   ByRef aR As Integer, ByRef mR As Integer, ByRef dR As Integer)
    dC = Day(fAte) - Day(fNac)
    mC = Month(fAte) - Month(fNac)
    aC = Year(fAte) - Year(fNac)
    If dC < 0 Then dC = dC + 30: mC = mC - 1
    If mC < 0 Then mC = mC + 12: aC = aC - 1
    
    aR = aC: mR = mC: dR = dC
    If semGest > 0 And semGest < 37 And aC < 2 Then
        Dim totalD As Long, diasP As Long, restoD As Long
        totalD = (CLng(aC) * 360) + (CLng(mC) * 30) + dC
        diasP = CLng(40 - semGest) * 7
        If totalD <= diasP Then
            aR = 0: mR = 0: dR = 0
        Else
            aR = aC - Int(diasP / 360)
            restoD = diasP Mod 360
            mR = mC - Int(restoD / 30)
            dR = dC - (restoD Mod 30)
            If dR < 0 Then dR = dR + 30: mR = mR - 1
            If mR < 0 Then mR = mR + 12: aR = aR - 1
        End If
    End If
End Sub

