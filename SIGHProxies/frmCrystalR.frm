VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCrystalR 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmCrystalR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CrvReportes 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "frmCrystalR.frx":0CCA
      DownPicture     =   "frmCrystalR.frx":118E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   3720
      Picture         =   "frmCrystalR.frx":167A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1365
   End
End
Attribute VB_Name = "frmCrystalR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Procesa y Muestra varios Reportes
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit



Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lc_Tabla As New ADODB.Recordset
Dim ln_Excel As Boolean
Dim ln_Archivo As String
Dim lc_TextoDelFiltro As String
Dim lcTitEESS As String, lcTitDireccion As String, lcTitTelefono As String
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes    'debb-27/05/2015


Property Let Tabla(lValue As ADODB.Recordset)
  Set lc_Tabla = lValue
End Property

Property Let Excel(lValue As Boolean)
  ln_Excel = lValue
End Property

Property Let Archivo(lValue As String)
  ln_Archivo = lValue
End Property

Private Sub btnCancelar_Click()
  Unload Me
End Sub

Property Let TextoDelFiltro(lValue As String)
  lc_TextoDelFiltro = lValue
End Property



Private Sub Form_Load()
  If Len(lc_TextoDelFiltro) > 250 Then
     lc_TextoDelFiltro = Left(lc_TextoDelFiltro, 250)
  End If

  Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
  Dim crParamDef As CRAXDRT.ParameterFieldDefinition

  On Error GoTo ErrHandler
  Screen.MousePointer = vbHourglass
    lcTitEESS = lcBuscaParametro.SeleccionaFilaParametro(205)
    lcTitDireccion = lcBuscaParametro.SeleccionaFilaParametro(206)
    lcTitTelefono = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)

  Set crReport = crApp.OpenReport(App.Path & "\plantillas\" & ln_Archivo & ".rpt", 1)
  Select Case ln_Archivo
  Case "EconConsumoXpto"
        Set crParamDefs = crReport.ParameterFields
        For Each crParamDef In crParamDefs
            Select Case crParamDef.ParameterFieldName
                Case "subTitulo"
                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
                Case "FechaHoraImpresion"
                    crParamDef.AddCurrentValue (lcBuscaParametro.RetornaFechaHoraServidorSQL)
                Case "lcEESS"
                    crParamDef.AddCurrentValue (lcTitEESS)
                Case "lcEESSdireccion"
                    crParamDef.AddCurrentValue (lcTitDireccion)
                Case "lcEESStelefono"
                    crParamDef.AddCurrentValue (lcTitTelefono)
            End Select
        Next
    Case "LabPorGrupoExamen"
        Set crParamDefs = crReport.ParameterFields
        For Each crParamDef In crParamDefs
            Select Case crParamDef.ParameterFieldName
                Case "subTitulo"
                    crParamDef.AddCurrentValue (lc_TextoDelFiltro)
               Case "lcEESS"
                   crParamDef.AddCurrentValue (lcTitEESS)
               Case "lcEESSdireccion"
                   crParamDef.AddCurrentValue (lcTitDireccion)
               Case "lcEESStelefono"
                   crParamDef.AddCurrentValue (lcTitTelefono)
            End Select
        Next
  Case Else
        Set crParamDefs = crReport.ParameterFields
        For Each crParamDef In crParamDefs
            Select Case crParamDef.ParameterFieldName
                Case "lcEESS"
                    crParamDef.AddCurrentValue (lcTitEESS)
                Case "lcEESSdireccion"
                    crParamDef.AddCurrentValue (lcTitDireccion)
                Case "lcEESStelefono"
                    crParamDef.AddCurrentValue (lcTitTelefono)
            End Select
        Next
  End Select
  crReport.Database.SetDataSource lc_Tabla
  If ln_Excel = True Then
  
      If lcBuscaParametro.SeleccionaFilaParametro(284) = "S" Then
          Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
          Select Case ln_Archivo
          Case "EconConsumoXpto"
              mo_ReglasReportes.ExportarRecordSetAexcel lc_Tabla, "EconConsumoXpto", lc_TextoDelFiltro, "", Me.hwnd
          Case "rPorGrupoExamen"
              mo_ReglasReportes.ExportarRecordSetAexcel lc_Tabla, "Por GRUPO DE EXAMEN", lc_TextoDelFiltro, "", Me.hwnd
          End Select
          Set mo_ReglasReportes = Nothing
    Else
        crReport.ExportOptions.DestinationType = crEDTDiskFile
        crReport.ExportOptions.FormatType = crEFTExcel70
        crReport.ExportOptions.DiskFileName = "c:\" & ln_Archivo & ".xls"
        crReport.Export (False)
        MsgBox "Se generó el archivo c:\" & ln_Archivo & ".xls"
    End If
  End If
  CrvReportes.ReportSource = crReport
  CrvReportes.ViewReport
  CrvReportes.Zoom 120
'
  mo_reglasComunes.grabaTablaAuditoria (crReport.Database.Tables.Item(1).Name & " " & _
                             Mid(lc_TextoDelFiltro, IIf(InStr(lc_TextoDelFiltro, "FILTROS: ") > 0, 10, 1)))   'debb-27/05/2015
  

  Screen.MousePointer = vbDefault
  Set crParamDefs = Nothing
  Set crParamDef = Nothing
  Set lc_Tabla = Nothing
  Screen.MousePointer = vbDefault
  Exit Sub
  
ErrHandler:
  If Err.Number = -2147206461 Then
    MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", vbCritical + vbOKOnly
  Else
    MsgBox Err.Description, vbCritical + vbOKOnly
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set crReport = Nothing
  Set crApp = Nothing
  Set lc_Tabla = Nothing
End Sub

Private Sub Form_Resize()
  CrvReportes.Top = 0
  CrvReportes.Left = 0
  CrvReportes.Height = ScaleHeight
  CrvReportes.Width = ScaleWidth
End Sub

