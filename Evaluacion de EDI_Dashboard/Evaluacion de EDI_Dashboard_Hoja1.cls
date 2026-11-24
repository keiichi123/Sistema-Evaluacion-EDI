VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim celdaF7 As Range
    Dim strVal As String
    Dim numVal As Double
    
    Set celdaF7 = Me.Range("F7")
    If Intersect(Target, Me.Range("F7:G7")) Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    On Error GoTo LimpiarSalida
    
    strVal = Trim(CStr(celdaF7.Value))
    
    ' Permitir celda vacía
    If strVal = "" Then GoTo LimpiarSalida
    
    ' Validar si es numérico, entero, y si es 0 o está entre 21 y 40
    If Not IsNumeric(strVal) Then
        celdaF7.ClearContents
        MsgBox "Dato rechazado: deje vacío, ingrese 0, o un número entero entre 21 y 40 semanas.", vbExclamation, "Edad Gestacional Inválida"
    Else
        numVal = CDbl(strVal)
        If Int(numVal) <> numVal Or (numVal <> 0 And (numVal < 21 Or numVal > 40)) Then
            celdaF7.ClearContents
            MsgBox "Dato rechazado: deje vacío, ingrese 0, o un número entero entre 21 y 40 semanas.", vbExclamation, "Edad Gestacional Inválida"
        End If
    End If

LimpiarSalida:
    Application.EnableEvents = True
End Sub
