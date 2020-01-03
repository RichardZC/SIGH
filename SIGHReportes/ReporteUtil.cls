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
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reportes excel
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Function NullToVacio(oValue As Variant) As String
    NullToVacio = IIf(IsNull(oValue), "", oValue)
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

    oWorkSheet.range(oWorkSheet.Cells(lFilaIni, lColumnaIni), oWorkSheet.Cells(lFilaFin, lColumnaFin)).Select
    
    With oExcelApp.Selection.borders(xlEdgeBottom)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.borders(xlEdgeTop)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.borders(xlEdgeRight)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.borders(xlEdgeLeft)
        .Weight = xlMedium
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.borders(xlInsideVertical)
        .Weight = xlThin
        .ColorIndex = 41
    End With
    With oExcelApp.Selection.borders(xlInsideHorizontal)
        .Weight = xlThin
        .ColorIndex = 41
    End With

End Sub

