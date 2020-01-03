Attribute VB_Name = "mdlLaboratorio"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Módulo para Laboratorio
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ExcelApplication As Excel.Application
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
'para tener la ventana activa
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Sub SeleccionaTexto(ByVal TextB As TextBox)
  'Selecciona el Contenido de una Caja de Texto
  TextB.SelStart = 0
  TextB.SelLength = Len(TextB.Text)
End Sub

Public Sub SeleccionaMask(ByVal MaskB As MaskEdBox)
  'Selecciona el Contenido de una Caja de Texto
  MaskB.SelStart = 0
  MaskB.SelLength = Len(MaskB.Text)
End Sub

Public Function Ubica_En_Combo(C As ComboBox, Co As String) As Integer
  Ubica_En_Combo = -1
  Dim Z, Y As Integer
  Z = C.ListCount
  For Y = 0 To Z - 1
    If C.List(Y) = Co Then
      Ubica_En_Combo = Y
      Exit For
    End If
  Next Y
End Function

Public Function EmpleadoTrabajaEnLaboratorio(idEmpleado As Long) As Boolean
  Dim rsTemp As New ADODB.Recordset
  Set rsTemp = mo_ReglasLaboratorio.LaboratorioSeleccionarRol(idEmpleado)
  If rsTemp.EOF = True And rsTemp.BOF = True Then
    EmpleadoTrabajaEnLaboratorio = False
  Else
    If rsTemp.RecordCount > 0 Then EmpleadoTrabajaEnLaboratorio = True
  End If
End Function

Function GalenhosExcelApplication() As Excel.Application
  If mo_ExcelApplication Is Nothing Then Set mo_ExcelApplication = New Excel.Application
  Set GalenhosExcelApplication = mo_ExcelApplication
End Function

Function GalenhosKillExcelApplication()
  Set mo_ExcelApplication = Nothing
End Function

