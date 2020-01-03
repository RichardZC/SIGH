VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHCpacienteHastaNanio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para lista de pacientes hasta NN años
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdTipoHistoria As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasArchivoClinico As New ReglasArchivoClinico
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_EdadMaxima As Long

Property Let EdadMaxima(lValue As Long)
    ml_EdadMaxima = lValue
End Property

Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdTipoHistoria(lValue As Long)
    ml_IdTipoHistoria = lValue
End Property
Property Let IdResponsable(lValue As Long)
    ml_IdResponsable = lValue
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New AHCpacientesMenoresNanio
    oFormulario.Show 1
End Sub


Sub CrearReporte(lbEnExcel As Boolean, lbElReporteEsEnAnios As Boolean, lcSector As String, lcComunidad As String, lnHwnd As Long)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    On Error GoTo ManejadorError
    'Filtra los Datos
    If lbElReporteEsEnAnios = True Then
        Set rsReporte = mo_ReglasArchivoClinico.HistoriasClinicasSeleccionarPorAnios(ml_EdadMaxima)
        If lbEnExcel = True Then
            mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Historias por Tipo", ml_TextoDelFiltro, "Nro historias: " & Trim(Str(rsReporte.RecordCount)), lnHwnd
        Else
            Set RpAHCconVIH.DataSource = rsReporte
            RpAHCconVIH.RightMargin = 10
            RpAHCconVIH.TopMargin = 10
            RpAHCconVIH.LeftMargin = 10
            RpAHCconVIH.BottomMargin = 10
            RpAHCconVIH.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            RpAHCconVIH.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            RpAHCconVIH.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            RpAHCconVIH.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
            RpAHCconVIH.Sections("cabecera").Controls("lblFiltro").Caption = ml_TextoDelFiltro
            Set RpAHCconVIH.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            RpAHCconVIH.Sections("pie").Controls("lblNhistorias").Caption = rsReporte.RecordCount
            RpAHCconVIH.Orientation = rptOrientLandscape
            RpAHCconVIH.Show 1
            'debb-27/05/2015
            
            mo_reglasComunes.grabaTablaAuditoria ("RpAHCconVIH: " & _
                                           Mid(ml_TextoDelFiltro, IIf(InStr(ml_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))

            '
        End If
   Else
      Set rsReporte = mo_ReglasArchivoClinico.PacientesSeleccionarPorSectorComunidad(lcSector, lcComunidad)
      If lbEnExcel = True Then
          mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Fichas Familiares por Sector/Comunidad", ml_TextoDelFiltro, "N°: " & Trim(Str(rsReporte.RecordCount)), lnHwnd
        Else
            Set RpAHCconVIH.DataSource = rsReporte
            RpAHCconVIH.RightMargin = 10
            RpAHCconVIH.TopMargin = 10
            RpAHCconVIH.LeftMargin = 10
            RpAHCconVIH.BottomMargin = 10
            RpAHCconVIH.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            RpAHCconVIH.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            RpAHCconVIH.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            RpAHCconVIH.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
            RpAHCconVIH.Sections("cabecera").Controls("Etiqueta2").Caption = "Lista de Fichas Familiares"
            RpAHCconVIH.Sections("cabecera").Controls("Etiqueta3").Caption = ""
            RpAHCconVIH.Sections("cabecera").Controls("lblFiltro").Caption = ml_TextoDelFiltro
            Set RpAHCconVIH.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            RpAHCconVIH.Sections("pie").Controls("lblNhistorias").Caption = rsReporte.RecordCount
            RpAHCconVIH.Orientation = rptOrientLandscape
            RpAHCconVIH.Show 1
            'debb-27/05/2015
            mo_reglasComunes.grabaTablaAuditoria ("RpAHCconVIH: " & _
                                           Mid(ml_TextoDelFiltro, IIf(InStr(ml_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))
            '
        End If
   End If
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Resume
    Exit Sub
End Sub




