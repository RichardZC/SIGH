VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clRptHistoriaClinicaHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idCuentaAtencion As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim oReglasCaja As New SIGHNegocios.ReglasCaja
Dim mo_ReporteUtil As New ReporteUtil

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property

Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property

Sub CrearReporteHistoriaClinicaDeLaAtencion(lcPlan As String, lcTransferencia As String, lcDx As String, lcCamaIng As String, lnUsuario As Long)
  Dim oExcel As Excel.Application
  Dim oWorkBookPlantilla As Workbook
  Dim oWorkBook As Workbook
  Dim oWorkSheet As Worksheet
  Dim rsReporte As New Recordset
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  On Error GoTo ManejadorError

  Set oExcel = GalenhosExcelApplication()  'New Excel.Application
    
  'Crea nueva hoja
  Set oWorkBook = oExcel.Workbooks.Add
  
  'Abre, copia y cierra la plantilla
  Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HHistoriaClinicaHosp.xls")
  oWorkBookPlantilla.Worksheets("HistoriaClinicaHosp").Copy Before:=oWorkBook.Sheets(1)
  oWorkBookPlantilla.Close
    
  'Activa la primera hoja
  Set oWorkSheet = oWorkBook.Sheets(1)
  'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
  oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
    
  Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
  If mo_AdminReportes.MensajeError <> "" Then
    MsgBox mo_AdminReportes.MensajeError, vbCritical, "Reporte"
  Else
    oWorkSheet.Cells(8, 85).Value = "'" & rsReporte!NroHistoriaClinica
    oWorkSheet.Cells(9, 85).Value = "'" & Trim(Str(ml_idCuentaAtencion))
    oWorkSheet.Cells(10, 33).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
    
    oWorkSheet.Cells(12, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad)
    oWorkSheet.Cells(14, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!EstadoCivil)
    oWorkSheet.Cells(15, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio)
    oWorkSheet.Cells(16, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio)
    oWorkSheet.Cells(17, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
    oWorkSheet.Cells(18, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio)
    oWorkSheet.Cells(19, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
    oWorkSheet.Cells(20, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
    oWorkSheet.Cells(21, 24).Value = "'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico))
    oWorkSheet.Cells(22, 24).Value = "'" & lcTransferencia
    '
    oWorkSheet.Cells(12, 81).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
    oWorkSheet.Cells(13, 24).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento)
    oWorkSheet.Cells(14, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Telefono)
    oWorkSheet.Cells(15, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NroDocumento) & " " & mo_ReporteUtil.NullToVacio(rsReporte!TipoDocumento)
    oWorkSheet.Cells(16, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio)
    oWorkSheet.Cells(17, 81).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio)
    oWorkSheet.Cells(18, 81).Value = "'" & lcCamaIng
    oWorkSheet.Cells(21, 81).Value = "'" & lcPlan
    '
    oWorkSheet.Cells(23, 28).Value = "'" & lcDx
    '
    oWorkSheet.Cells(53, 4).Value = "'Usuario: " & oReglasCaja.SeleccionaDatosCajero(lnUsuario, sghApellidosYnombres)
    '
    If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
      oExcel.Visible = True
      oWorkSheet.PrintPreview
    Else
      oWorkSheet.PageSetup.PrintArea = "$A$1:$DB$54"
      oWorkSheet.PrintOut
    End If
  End If
    
  oWorkBook.Close SaveChanges:=False

Exit Sub
ManejadorError:
  Select Case Err.Number
    Case 1004
      MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
      MsgBox Err.Description
  End Select
  Exit Sub
End Sub
