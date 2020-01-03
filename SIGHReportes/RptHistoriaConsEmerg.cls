VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHistoriaConsEmerg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idCuentaAtencion As Long
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_ReporteUtil As New ReporteUtil
Dim mo_DOAtencionesCE As New DOAtencionesCE
Dim ml_Plan As String
Property Let Plan(lValue As String)
    ml_Plan = lValue
End Property
Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property
Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property
Sub CrearReporteHistoriaClinicaConsultorioEmerg()
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
    Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\EHistoriaConsultorioEmerg.xls")
    oWorkBookPlantilla.Worksheets("HistoriaConsultorioEmerg").Copy Before:=oWorkBook.Sheets(1)
    oWorkBookPlantilla.Close
    
    'Activa la primera hoja
    Set oWorkSheet = oWorkBook.Sheets(1)
    'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
    oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
    
    'Izquierdo
    CargaAtencionCEJamo
    Set rsReporte = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(ml_idAtencion)
    oWorkSheet.Cells(6, 29).Value = "'" & rsReporte!NroHistoriaClinica
    oWorkSheet.Cells(7, 29).Value = "'" & mo_ReporteUtil.ArmarNombreDePaciente(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaterno), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaterno), mo_ReporteUtil.NullToVacio(rsReporte!PrimerNombre), mo_ReporteUtil.NullToVacio(rsReporte!SegundoNombre))
    oWorkSheet.Cells(8, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaIngreso)
    oWorkSheet.Cells(9, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!HoraIngreso)
    oWorkSheet.Cells(10, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Servicio)
    oWorkSheet.Cells(11, 29).Value = UCase("'" & mo_ReporteUtil.ArmarNombreDeMedico(mo_ReporteUtil.NullToVacio(rsReporte!ApellidoPaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!ApellidoMaternoMedico), mo_ReporteUtil.NullToVacio(rsReporte!NombresMedico)))
    oWorkSheet.Cells(12, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Telefono)
    oWorkSheet.Cells(13, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DireccionDomicilio)
    oWorkSheet.Cells(14, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DepartamentoDomicilio)
    oWorkSheet.Cells(15, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!DistritoDomicilio)
    oWorkSheet.Cells(16, 29).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!NombreAcompaniante)
    
    'Derecho
    oWorkSheet.Cells(6, 80).Value = "'" & Trim(Str(ml_idCuentaAtencion)) & " " & ml_Plan
    oWorkSheet.Cells(8, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!Edad)
    oWorkSheet.Cells(9, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!FechaNacimiento)
    oWorkSheet.Cells(10, 85).Value = "'" & IIf(mo_ReporteUtil.NullToVacio(rsReporte!idTipoSexo) = "1", "Masculino", "Femenino")
    oWorkSheet.Cells(14, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!ProvinciaDomicilio)
    oWorkSheet.Cells(15, 85).Value = "'" & mo_ReporteUtil.NullToVacio(rsReporte!CentroPobladoDomicilio)
    If mo_DOAtencionesCE.idAtencion > 0 Then
       oWorkSheet.Cells(23, 4).Value = mo_DOAtencionesCE.TriajePresion
       oWorkSheet.Cells(23, 25).Value = "'" & mo_DOAtencionesCE.TriajeFrecRespiratoria
       oWorkSheet.Cells(23, 50).Value = "'" & mo_DOAtencionesCE.TriajePulso
       oWorkSheet.Cells(23, 73).Value = "'" & mo_DOAtencionesCE.TriajeTemperatura
       'oWorkSheet.Cells(23, 85).Value=
    End If
    If lcBuscaParametro.SeleccionaFilaParametro(216) <> "1" Then
        oExcel.Visible = True
        oWorkSheet.PrintPreview
    Else
        oWorkSheet.PrintOut
        oWorkBook.Close SaveChanges:=False
    End If
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




Sub CargaAtencionCEJamo()
            On Error GoTo ErrJamo
            Dim oConexion As New Connection
            Dim lcBuscaParametro As New SIGHDatos.Parametros
            Dim oAtencionesCE As New AtencionesCE
            oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            mo_DOAtencionesCE.idAtencion = ml_idAtencion
            Set oAtencionesCE.Conexion = oConexion
            If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
                mo_DOAtencionesCE.idAtencion = 0
                Exit Sub
            End If
            oConexion.Close
            Set oConexion = Nothing
            Set oAtencionesCE = Nothing
            Exit Sub
ErrJamo:
End Sub

