VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReporteUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Reportes Excel
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Function NullToVacio(oValue As Variant) As String
    NullToVacio = IIf(IsNull(oValue), "", oValue)
End Function
Function DevuelveEpsPorcentaje(oValue As Variant) As Double
        DevuelveEpsPorcentaje = 0
        If Not IsNull(oValue) Then
           If oValue > 0 Then
              DevuelveEpsPorcentaje = oValue
           End If
        End If
End Function
Function DevuelveEPScubre(lnEpsPorcentaje As Double) As String
    DevuelveEPScubre = IIf(lnEpsPorcentaje > 0, " (cubre: " & Trim(str(lnEpsPorcentaje)) & "% )", "")
End Function
Function ArmarNombreDePaciente(sApellidoPaterno As String, sApellidoMaterno As String, sPrimerNombre As String, sSegundoNombre As String) As String
    ArmarNombreDePaciente = sApellidoPaterno + " " + sApellidoMaterno + " " + sPrimerNombre + " " + sSegundoNombre
End Function
Function ArmarNombreDeMedico(sApellidoPaterno As String, sApellidoMaterno As String, sNombres As String) As String
    ArmarNombreDeMedico = sApellidoPaterno + " " + sApellidoMaterno + " " + sNombres
End Function
Function ArmarNombreDeEmpleado(sApellidoPaterno As String, sApellidoMaterno As String, sNombres As String) As String
    ArmarNombreDeEmpleado = sApellidoPaterno + " " + sApellidoMaterno + " " + sNombres
End Function

Function ArmarDireccionDePaciente(sDireccion As String, sNumero As String, sManzana As String, sLote As String, sSector As String, sEtapa As String) As String
    ArmarDireccionDePaciente = sDireccion + " " + _
                            IIf(sNumero = "", "", "Nº " + sNumero) + " " + _
                            IIf(sManzana = "", "", "Mz: " + sManzana) + " " + _
                            IIf(sLote = "", "", "Lt: " + sLote) + " " + _
                            IIf(sSector = "", "", "Sec: " + sSector) + " " + _
                            IIf(sEtapa = "", "", "Etapa: " + sEtapa)
End Function

Sub ExcelCuadricularRango(oExcelApp As Excel.Application, oWorkSheet As Worksheet, lFilaIni As Long, lColumnaIni As Integer, lFilaFin As Long, lColumnaFin As Integer)
On Error Resume Next

    oWorkSheet.Range(oWorkSheet.Cells(lFilaIni, lColumnaIni), oWorkSheet.Cells(lFilaFin, lColumnaFin)).Select
    
    With oExcelApp.Selection.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.Borders(xlEdgeTop)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.Borders(xlEdgeRight)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.Borders(xlEdgeLeft)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.Borders(xlInsideVertical)
        .Weight = xlThin
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.Borders(xlInsideHorizontal)
        .Weight = xlThin
        .ColorIndex = 41
    End With

End Sub

Sub ExcelOpenOfficeCuadricularRango(oPlage As Object, lnGrosorLinea As Integer)
Dim Style As Object
Dim Border As Object

On Error Resume Next

    Set Border = oPlage.TableBorder
   
    Set Style = Border.TopLine
    Style.OuterLineWidth = lnGrosorLinea
    Border.TopLine = Style
    
    Set Style = Border.bottomline
    Style.OuterLineWidth = lnGrosorLinea
    Border.bottomline = Style
    
    Set Style = Border.leftline
    Style.OuterLineWidth = lnGrosorLinea
    Border.leftline = Style
    
    Set Style = Border.rightline
    Style.OuterLineWidth = lnGrosorLinea
    Border.rightline = Style
    
    oPlage.TableBorder = Border

End Sub

