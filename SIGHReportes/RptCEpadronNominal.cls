VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCEpadronNominal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para emisión del Formato HIS
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Sub EjecutaFormulario()
    Dim oFormulario As New CEpadronNominal
    oFormulario.Show 1
End Sub

Sub CrearReporte(ldFechaInicial As Date, ldFechaFinal As Date, lnIdTipoSexo As Long, lnEdadenAnios As Long, _
                 lnIdPaciente As Long, lbEnExcel As Boolean, ml_TextoDelFiltro As String, lnHwnd As Long)

        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
        Dim rsReporte As New Recordset
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Set rsReporte = PacientesMovimientosFiltro(ldFechaInicial, ldFechaFinal, lnEdadenAnios, _
                                                   lnIdPaciente)
        If lnIdTipoSexo > 0 Then
           rsReporte.Filter = "idTipoSexo=" & lnIdTipoSexo
        End If
        If rsReporte.RecordCount = 0 Then
            MsgBox "No se encontró información con esos datos", vbInformation, ""
        Else
            If lbEnExcel = True Then
                mo_ReglasReportes.ExportarRecordSetAexcel rsReporte, "Padrón Nominal", ml_TextoDelFiltro, _
                                              "Nro Cuentas: " & Trim(Str(rsReporte.RecordCount)), lnHwnd
            Else
                Set RpCEpadronNominal.DataSource = rsReporte
                RpCEpadronNominal.RightMargin = 10
                RpCEpadronNominal.TopMargin = 10
                RpCEpadronNominal.LeftMargin = 10
                RpCEpadronNominal.BottomMargin = 10
                RpCEpadronNominal.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
                RpCEpadronNominal.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
                RpCEpadronNominal.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
                RpCEpadronNominal.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
                RpCEpadronNominal.Sections("cabecera").Controls("lblFiltro").Caption = ml_TextoDelFiltro
                Set RpCEpadronNominal.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                RpCEpadronNominal.Sections("pie").Controls("lblNhistorias").Caption = rsReporte.RecordCount
                RpCEpadronNominal.Orientation = rptOrientPortrait
                RpCEpadronNominal.Show 1
                'debb-27/05/2015
                Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
                mo_reglasComunes.grabaTablaAuditoria ("RpCEpadronNominal: " & _
                                               Mid(ml_TextoDelFiltro, IIf(InStr(ml_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))
                Set mo_reglasComunes = Nothing
                '
            End If
        End If
        Set mo_ReglasReportes = Nothing
        Set rsReporte = Nothing
        Set lcBuscaParametro = Nothing
End Sub


Function PacientesMovimientosFiltro(ldFechaInicio As Date, ldFechaFin As Date, lnEdadAnios As Long, lnIdPaciente As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PacientesMovimientosFiltro"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EdadAnios", adInteger, adParamInput, 0, lnEdadAnios): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set PacientesMovimientosFiltro = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
