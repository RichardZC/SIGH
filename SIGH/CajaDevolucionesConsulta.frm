VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{1416D7C5-8A28-11CF-9236-444553540000}#8.0#0"; "PVXPLORE8.ocx"
Begin VB.Form PrincipalC1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Información para Clínica"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11445
   Icon            =   "CajaDevolucionesConsulta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PVExplorerLib.PVExplorer PVExplorer1 
      Height          =   5475
      Left            =   15
      TabIndex        =   0
      Top             =   1320
      Width           =   11280
      _Version        =   524288
      Indentation     =   0
      SourceChannel1  =   ""
      TargetChannel1  =   ""
      PathSeparator   =   ""
      Image1          =   "CajaDevolucionesConsulta.frx":0152
      SourceChannel2  =   ""
      TargetChannel2  =   ""
      Image2          =   "CajaDevolucionesConsulta.frx":0ED0
      Image3          =   "CajaDevolucionesConsulta.frx":2DCE
      CheckBoxes2     =   -1  'True
      FileName        =   ""
      DataMember      =   ""
      DataField0      =   ""
      DataField1      =   ""
      DataField2      =   ""
      DataField3      =   ""
      DataField4      =   ""
      DataField5      =   ""
      DataField6      =   ""
      DataField7      =   ""
      DataField8      =   ""
      DataField9      =   ""
      DataField10     =   ""
      DataField11     =   ""
      DataField12     =   ""
      DataField13     =   ""
      DataField14     =   ""
      DataField15     =   ""
      DataField16     =   ""
      DataField17     =   ""
      DataField18     =   ""
      DataField19     =   ""
      PaneDisplay     =   1
      CaptionMode     =   2
      DynamicResize   =   -1  'True
      BorderWidth     =   25
      _ExtentX        =   19897
      _ExtentY        =   9657
      _StockProps     =   70
   End
   Begin VB.Frame Frame 
      BackColor       =   &H8000000D&
      Height          =   1170
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   11280
      Begin VB.Label lblHospital 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clínica"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblPcServidor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pc y Servidor"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   465
         Width           =   945
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ColumnsImageList 
      Left            =   6840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":3B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":3C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":3E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":3F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":4010
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CajaDevolucionesConsulta.frx":4104
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivos 
      Caption         =   "Archivos"
      Begin VB.Menu mnuArchivosArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArchivosGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuCE 
      Caption         =   "Consulta Externa"
      Begin VB.Menu mnuCEarray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCEguiones 
         Caption         =   "---"
      End
   End
   Begin VB.Menu mnuHospitalizacion 
      Caption         =   "Hospitalización"
      Begin VB.Menu mnuHospArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHospGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuEmergencia 
      Caption         =   "Emergencia"
      Begin VB.Menu mnuEmergArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEmergGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuFarmacia 
      Caption         =   "Farmacia"
      Begin VB.Menu mnuFarmaciaArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFarmGuiones 
         Caption         =   "---"
      End
   End
   Begin VB.Menu mnuLaboratorio 
      Caption         =   "Laboratorio"
      Begin VB.Menu mnuLaboArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLaboGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuImagenes 
      Caption         =   "Imagenes"
      Begin VB.Menu mnuImagArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImagGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuContabilidad 
      Caption         =   "Contabilidad"
      Begin VB.Menu mnuContArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuOtros 
      Caption         =   "Otros"
      Begin VB.Menu mnuOtrosArray 
         Caption         =   "NoVisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOtrosGuiones 
         Caption         =   "----"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "PrincipalC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Ordenes de Farmacia y Servicio sin pago en CAJA
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim lc_NombrePc As String
Dim mrs_Reportes As New Recordset
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Const lcXArchivos As String = "ARCHIVO CLINICO"
Const lcXce As String = "CONSULTA EXTERNA"
Const lcXECONOMIA As String = "ECONOMIA"
Const lcXEMERGENCIA As String = "EMERGENCIA"
Const lcXFARMACIA As String = "FARMACIA"
Const lcXHERRAMIENTAS As String = "HERRAMIENTAS"
Const lcXHOSPITALIZACION As String = "HOSPITALIZACION"
Const lcXIMAGENOLOGIA As String = "IMAGENOLOGIA"
Const lcXLABORATORIO As String = "LABORATORIO"


Private Sub LlenaOpcionesYsubOpciones()
    Dim DefaultColumns As pvxColumnHeaders
    Dim CustomColumns1 As pvxColumnHeaders
    Dim CustomColumns2 As pvxColumnHeaders
    Dim Column As pvxColumnHeader
    Dim Nodes As pvxNodes
    Dim Node As pvxNode
    Dim Entry As pvxNode
    Dim Item As pvxListItem

    PVExplorer1.LeftPaneWidth = 180
    PVExplorer1.ListView.View = pvxReport
    Set DefaultColumns = PVExplorer1.ListView.ColumnHeaders
    
    PVExplorer1.ToolBar.ImageList = ColumnsImageList
    PVExplorer1.ToolBar.CreateIEToolBar
    'DefaultColumns.ImageList = ColumnsImageList

    Set CustomColumns1 = PVExplorer1.ListView.CreateColumnHeaders
    Set CustomColumns2 = PVExplorer1.ListView.CreateColumnHeaders
    
    ' Add column headers for the ListView pane in pvxReport View
    Set Column = DefaultColumns.Add(0, "Last Name", 100, 0)
    Column.Image = 0
    Set Column = DefaultColumns.Add(1, "First Name", 140, 0)
    Column.Image = 1
    Set Column = DefaultColumns.Add(2, "Product Number", 100, 0)
    Column.Image = 2
    Set Column = DefaultColumns.Add(3, "Address", 160, 0)
    Column.Image = 3
    Set Column = DefaultColumns.Add(4, "City", 100, 0)
    Column.Image = 4
    Set Column = DefaultColumns.Add(5, "State", 20, 0)
    Column.Image = 5
    Set Column = DefaultColumns.Add(6, "Zip Code", 40, 0)
    Column.Image = 6

    CustomColumns1.ImageList = ColumnsImageList
    Set Column = CustomColumns1.Add(0, "Product Name", 100, 0)
    Column.Image = 4
    Set Column = CustomColumns1.Add(1, "Product ID", 500, 0)
    Column.Image = 2
    Set Column = CustomColumns1.Add(2, "Price", 100, 0)
    Set Column = CustomColumns1.Add(3, "Quantity", 100, 0)
    Column.Image = 1
    
    Set Column = CustomColumns2.Add(0, "Opciones", 500, 0)
    Set Column = CustomColumns2.Add(1, "Product code", 1, 0)
    
    
    
    ' Add a series of Node objects to the TreeView Nodes collection
    Set Nodes = PVExplorer1.TreeView.Nodes
    
    Set Node = Nodes.Add(Nothing, pvxLast, "Lista de Módulos", 0, 1)
    Node.ViewerType = pvxNoPane
    Node.BoldText = True

    Dim rsItems As New Recordset
    Dim rsGrupos As New Recordset
    Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim lnFila As Integer
    Set rsItems = mo_AdminSeguridad.RolesItemsSeleccionarItemsPorUsuarioYGrupoSql2000(sighentidades.Usuario, 0)
    Set rsGrupos = mo_AdminSeguridad.RolesItemsSeleccionarGruposPorUsuarioSql2000(sighentidades.Usuario)
    Do While Not rsGrupos.EOF
'            Grupo.Key = rsGrupos!Clave
'            Grupo.Caption = rsGrupos!Texto
            Set Entry = Nodes.Add(Node, pvxLast, rsGrupos!Texto, 0, 1)
            Set Entry.ListItems.ColumnHeaders = CustomColumns2
            '
            
            '
            Set rsItems = mo_AdminSeguridad.RolesItemsSeleccionarItemsPorUsuarioYGrupoSql2000(sighentidades.Usuario, rsGrupos!IdListGrupo)
            rsItems.Filter = "IdListGrupo=" & rsGrupos!IdListGrupo
            '
            lnFila = 1
            Do While Not rsItems.EOF
                Set Item = Entry.ListItems.Add(0, rsItems!Texto, lnFila, lnFila)
                Item.SubItems(1) = rsItems!Clave
'                ListItem.Key = RsItems!Clave
'                ListItem.Text = RsItems!Texto
                '
                lnFila = lnFila + 1
                rsItems.MoveNext
            Loop
            rsItems.Close
            '
            rsGrupos.MoveNext
    Loop
    Set rsItems = Nothing
    Set rsGrupos = Nothing
    Set mo_AdminSeguridad = Nothing
    Node.Expanded = True
End Sub

Sub CargaDatosGenerales()
    Dim oReglasCaja As New SIGHNegocios.ReglasCaja
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Me.lblPcServidor.Caption = "Pc/Servidor:" & sighentidades.RetornaNombrePC & "/" & lcBuscaParametro.RetornaNombreDeServidor
    Me.lblUsuario.Caption = "Usuario: " & oReglasCaja.SeleccionaDatosCajero(sighentidades.Usuario, sghUsuario)
    Me.lblHospital.Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
    ml_IdUsuarioAuditoria = sighentidades.Usuario
    Set oReglasCaja = Nothing
    Set lcBuscaParametro = Nothing
   
End Sub

Private Sub Form_Load()
    lc_NombrePc = sighentidades.RetornaNombrePC
    CargaDatosGenerales
    PVExplorer1.CaptionMode = pvxNoCaption
    LlenaOpcionesYsubOpciones
    LlenaReportes
End Sub

Private Sub Form_Resize()
'    PVExplorer1.Top = Me.Top
'    PVExplorer1.Left = Me.Left
'    PVExplorer1.Height = Me.Height
'    PVExplorer1.Width = Me.Width
End Sub




Private Sub mnuCEarray_Click(Index As Integer)
    MuestraReporte lcXce, Index
End Sub

Private Sub mnuContArray_Click(Index As Integer)
MuestraReporte lcXECONOMIA, Index
End Sub

Private Sub mnuEmergArray_Click(Index As Integer)
   MuestraReporte lcXEMERGENCIA, Index
End Sub

Private Sub mnuFarmaciaArray_Click(Index As Integer)
    MuestraReporte lcXFARMACIA, Index
End Sub

Private Sub mnuHospArray_Click(Index As Integer)
    MuestraReporte lcXHOSPITALIZACION, Index
End Sub

Private Sub mnuImagArray_Click(Index As Integer)
    MuestraReporte lcXIMAGENOLOGIA, Index
End Sub

Private Sub mnuLaboArray_Click(Index As Integer)
    MuestraReporte lcXLABORATORIO, Index
End Sub

Private Sub mnuOtrosArray_Click(Index As Integer)
    MuestraReporte lcXHERRAMIENTAS, Index
End Sub

Private Sub mnuSalir_Click()
    End
End Sub



Sub LlenaReportes()
    With mrs_Reportes
         .Fields.Append "id_menuReporte", adVarChar, 100
         .Fields.Append "Reporte", adVarChar, 200
         .Fields.Append "Modulo", adVarChar, 100
         .Fields.Append "OpcionMenu", adInteger
         .LockType = adLockOptimistic
         .Open
    End With
    Dim oRsTmp As New Recordset
    Dim lcSql As String, lcIdModulo As String, lnOpcionMenu As Integer, lcReporte As String
    Set oRsTmp = mo_AdminSeguridad.RetornaOpcionesReporteQueTieneAcceso(sighentidades.Usuario)
    oRsTmp.Filter = "Modulo<>null"
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.Sort = "modulo,reporte"
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          lcIdModulo = oRsTmp!Modulo
          Do While Not oRsTmp.EOF And lcIdModulo = oRsTmp!Modulo
                lcReporte = oRsTmp!reporte
                If Left(lcReporte, 4) = "----" Then
                    oRsTmp.MoveNext
                Else
                    Select Case oRsTmp!Modulo
                    Case lcXArchivos
                         lnOpcionMenu = Agregar(lcReporte, mnuArchivosArray)
                    Case lcXce
                         lnOpcionMenu = Agregar(lcReporte, mnuCEarray)
                    Case lcXECONOMIA
                         lnOpcionMenu = Agregar(lcReporte, mnuContArray)
                    Case lcXEMERGENCIA
                         lnOpcionMenu = Agregar(lcReporte, mnuEmergArray)
                    Case lcXFARMACIA
                         lnOpcionMenu = Agregar(lcReporte, mnuFarmaciaArray)
                    Case lcXHERRAMIENTAS
                         lnOpcionMenu = Agregar(lcReporte, mnuOtrosArray)
                    Case lcXHOSPITALIZACION
                         lnOpcionMenu = Agregar(lcReporte, mnuHospArray)
                    Case lcXIMAGENOLOGIA
                         lnOpcionMenu = Agregar(lcReporte, mnuImagArray)
                    Case lcXLABORATORIO
                         lnOpcionMenu = Agregar(lcReporte, mnuLaboArray)
                    End Select
                    mrs_Reportes.AddNew
                    mrs_Reportes!id_menuReporte = oRsTmp!id_menuReporte
                    mrs_Reportes!reporte = lcReporte
                    mrs_Reportes!Modulo = oRsTmp!Modulo
                    mrs_Reportes!OpcionMenu = lnOpcionMenu
                    mrs_Reportes.Update
                    Do While Not oRsTmp.EOF And lcIdModulo = oRsTmp!Modulo And lcReporte = oRsTmp!reporte
                       oRsTmp.MoveNext
                       If oRsTmp.EOF Then
                          Exit Do
                       End If
                    Loop
                End If
                If oRsTmp.EOF Then
                   Exit Do
                End If
          Loop
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub
Function Agregar(TextoDeMenu As String, QueMenu As Object) As Integer
    Dim indice As Integer
    indice = QueMenu.Count
    Load QueMenu(indice)
    QueMenu(indice).Caption = TextoDeMenu
    QueMenu(indice).Visible = True
    'QueMenu(indice).BackColor = vbBlue
    Agregar = indice
End Function
Private Sub mnuArchivosArray_Click(Index As Integer)
       MuestraReporte lcXArchivos, Index
End Sub

Sub MuestraReporte(lcModuloReporte As String, lnIndex As Integer)
    mrs_Reportes.Filter = "modulo='" & lcModuloReporte & "' and opcionMenu=" & lnIndex
    If mrs_Reportes.RecordCount > 0 Then
        Select Case lcModuloReporte
        Case lcXArchivos
             Select Case mrs_Reportes!reporte
             Case "Historias Clínicas con devolución pendiente al Archivo Clínico"
                    Dim oRptHCnoLlegaAC As New SIGHReportes.RptAHhcNoLlegaAC
                    oRptHCnoLlegaAC.EjecutaFormulario
                    Set oRptHCnoLlegaAC = Nothing
             Case "Historias Clínicas con problemas"
                   Dim oRptHCnoUsadas As New SIGHReportes.RptAHhcNOusadas
                   oRptHCnoUsadas.EjecutaFormulario
                   Set oRptHCnoUsadas = Nothing
             Case "Historias clinicas por tipo de historia"
                    Dim oRpt222 As New SIGHReportes.RptAHCconVIH
                    oRpt222.EjecutaFormulario
                    Set oRpt222 = Nothing
             Case "Historias Clínicas solicitadas por Médico"
                    Dim oRpt220 As New SIGHReportes.RptAHSolicPorMedico
                    oRpt220.EjecutaFormulario
                    Set oRpt220 = Nothing
             Case "Historias Clínicas solicitadas por Servicio"
                    Dim oRpt219 As New SIGHReportes.RptAHSolicPorServ
                    oRpt219.EjecutaFormulario
                    Set oRpt219 = Nothing
             Case "Historias Solicitadas Por Servicio"
                    Dim oSolicitud As New SIGHReportes.clSolicitudHistorias
                    oSolicitud.TipoReporte = "RPT_HISTORIAS_SERVICIO"
                    oSolicitud.idUsuario = ml_IdUsuarioAuditoria
                    oSolicitud.EjecutaFormulario
                    Set oSolicitud = Nothing
             Case "Lista de Historias para enviar al ARCHIVO PASIVO"
                    Dim oRptAHCMovimFormatos As New SIGHReportes.RptAHCMovimFormatos
                    oRptAHCMovimFormatos.EjecutaFormulario
                    Set oRptAHCMovimFormatos = Nothing
             Case "Movimientos de Historias"
                    Dim oRptAHCMovimEntSal As New SIGHReportes.RptAHCMovimEntSal
                    oRptAHCMovimEntSal.EjecutaFormulario
                    Set oRptAHCMovimEntSal = Nothing
             Case "Pacientes menores a N años"
                    Dim oRptMovimientoHistorias As New SIGHReportes.RptAHCpacienteHastaNanio
                    oRptMovimientoHistorias.EjecutaFormulario
                    Set oRptMovimientoHistorias = Nothing
             Case "Relación de Historias Clínicas de Pacientes Judiciales"
                    Dim oRpt223 As New SIGHReportes.RptAHSolicPorTipo
                    oRpt223.EjecutaFormulario
                    Set oRpt223 = Nothing
             Case "Reporte Historias Solicitadas Por Médico"
                    Dim oSolicitudMedico As New SIGHReportes.clSolicitudHistorias
                    oSolicitudMedico.TipoReporte = "RPT_HISTORIAS_MEDICO"
                    oSolicitudMedico.idUsuario = ml_IdUsuarioAuditoria
                    oSolicitudMedico.EjecutaFormulario
                    Set oSolicitudMedico = Nothing
             End Select
        Case lcXce
             Select Case mrs_Reportes!reporte
             Case "Citados y/o atendidos x Consultorios"
                    Dim oRptHosp2 As New SIGHProxies.clReportesEgreHosp
                    oRptHosp2.IdTipoReporte = sighentidades.sghReporteEgresosHospitalario
                    oRptHosp2.idTipoServicio = 2
                    oRptHosp2.EjecutaFormulario
                    Set oRptHosp2 = Nothing
             Case "Cupos Asignados"
                    Dim oRptCuposAsignados As New SIGHReportes.clCuposAsignadosRep
                    oRptCuposAsignados.EjecutaFormulario
                    Set oRptCuposAsignados = Nothing
             Case "Frecuencia de Dx de Pacientes atendidos"
                    Dim oRpt236 As New SIGHReportes.RptCEdx
                    oRpt236.EjecutaFormulario
                    Set oRpt236 = Nothing
             Case "Frecuencia de GASTOS DE SERVICIOS de Pacientes"
                    Dim oRpt237 As New SIGHReportes.RptCEservi
                    oRpt237.EjecutaFormulario
                    Set oRpt237 = Nothing
             Case "Imprime Formato HIS"
                    Dim oRpt234 As New SIGHProxies.RptCEhis
                    oRpt234.EjecutaFormulario
                    Set oRpt234 = Nothing
             Case "Indicador de Atenciones vs Atendidos"
                    Dim oRpt238 As New SIGHReportes.RptCEatenciones
                    oRpt238.EjecutaFormulario
                    Set oRpt238 = Nothing
             Case "Morbilidad Frecuente"
                    Dim oRptMorbilidadCE As New SIGHReportes.RptHMorbCE
                    oRptMorbilidadCE.EjecutaFormulario
                    Set oRptMorbilidadCE = Nothing
             Case "Padron Nominal"
                    Dim oRptCEpadronNominal As New RptCEpadronNominal                   'debb-2/3/2015
                    oRptCEpadronNominal.EjecutaFormulario                               'debb-2/3/2015
                    Set oRptCEpadronNominal = Nothing                                    'debb-2/3/2015
             Case "Reporte de programación médica"
                    Dim oProgMedicaRpt As New SIGHReportes.clProgramMedica
                    oProgMedicaRpt.EjecutaFormulario
                    Set oProgMedicaRpt = Nothing
             Case "Reportes para el módulo Materno"
                    Dim oRptRepMaterno As New SIGHReportes.clCeMaterno
                    oRptRepMaterno.EjecutaFormulario
                    Set oRptRepMaterno = Nothing
             Case "Reportes para el módulo Niño Sano - Indicadores"
                    Dim oRptRepPerinatalIndicadores As New SIGHReportes.clCePerinatalIndicadores
                    oRptRepPerinatalIndicadores.EjecutaFormulario
                    Set oRptRepPerinatalIndicadores = Nothing
             Case "Reportes para el módulo Perinatal"
                    Dim oRptRepPerinatal As New SIGHReportes.clCePerinatal
                    oRptRepPerinatal.EjecutaFormulario
                    Set oRptRepPerinatal = Nothing
             End Select
        Case lcXHOSPITALIZACION
             Select Case mrs_Reportes!reporte
             Case "Camas Hospitalarias por Dpto/Servicio"
                    Dim oRpt216 As New SIGHReportes.RptHCamas
                    oRpt216.EjecutaFormulario
                    Set oRpt216 = Nothing
             Case "Días cama Hospitalaria por Dpto/Servicio"
                    Dim oRpt217 As New SIGHReportes.RptHCamaDias
                    oRpt217.EjecutaFormulario
                    Set oRpt217 = Nothing
             Case "Días de Estancia Hospitalaria por Dpto/Servicios"
                    Dim oRpt214 As New SIGHReportes.RptHEstanciaH
                    oRpt214.EjecutaFormulario
                    Set oRpt214 = Nothing
             Case "Días Paciente Hospitalario por Dpto/Servicio"
                    Dim oRpt218 As New SIGHReportes.RptHDiasPaciente
                    oRpt218.EjecutaFormulario
                    Set oRpt218 = Nothing
             Case "Egresos Hospitalarios (Epicrisis)"
                    Dim oRptHosp As New SIGHProxies.clReportesEgreHosp
                    oRptHosp.IdTipoReporte = sighentidades.sghReporteEgresosHospitalario
                    oRptHosp.idTipoServicio = 0
                    oRptHosp.EjecutaFormulario
                    Set oRptHosp = Nothing
             Case "Egresos Hospitalarios por Dpto/Servicios"
                    Dim oRpt24 As New SIGHReportes.RptHEgresosHosp
                    oRpt24.EjecutaFormulario
                    Set oRpt24 = Nothing
             Case "Indicador Promedio Permanencia por Dpto/Servicio"
                    Dim oRpt215 As New SIGHReportes.RptHPrPermanencia
                    oRpt215.EjecutaFormulario
                    Set oRpt215 = Nothing
             Case "Indicadores Hospitalarios por Dpto/Servicio/Especialidad"
                    Dim oRpt13 As New SIGHReportes.RptHIndicadorAnual
                    oRpt13.EjecutaFormulario
                    Set oRpt13 = Nothing
             Case "Indicadores Hospitalarios por Meses"
                    Dim oRpt22 As New SIGHReportes.RptHIndicadorMeses
                    oRpt22.EjecutaFormulario
                    Set oRpt22 = Nothing
             Case "Ingresos Hospitalarios"
                    Dim oRptIngHosp As New SIGHProxies.clReporteIngrHosp
                    oRptIngHosp.IdTipoReporte = sighentidades.sghReporteIngresosHospitalario
                    oRptIngHosp.EjecutaFormulario
                    Set oRptIngHosp = Nothing
             Case "Ingresos Hospitalarios por Dpto/Servicios"
                    Dim oRpt25 As New SIGHReportes.RptHIngresosHosp
                    oRpt25.EjecutaFormulario
                    Set oRpt25 = Nothing
             Case "Mortalidad Hospitalaria por causa básica, segun ciclos de vida Dpto/Especialidad"
                    Dim oRpt29 As New SIGHReportes.RptHMortalidad
                    oRpt29.EjecutaFormulario
                    Set oRpt29 = Nothing
             Case "Primeras causas de morbilidad Hospitalaria por Diagnósticos, según ciclos de vida por Dpto/Especialidad"
                    Dim oRpt212 As New SIGHReportes.RptHMorbilidad
                    oRpt212.EjecutaFormulario
                    Set oRpt212 = Nothing
             Case "Reporte de Censo Hospitalario"
                    Dim oRptCensoHospitalario As New SIGHReportes.clAtencionesCenso
                    oRptCensoHospitalario.EjecutaFormulario
                    Set oRptCensoHospitalario = Nothing
             Case "Reporte de Procedimientos Hospitalarios por Dpto/Especialidad"
                    Dim oRpt213 As New SIGHReportes.RptHProcedimientos
                    oRpt213.EjecutaFormulario
                    Set oRpt213 = Nothing
             Case "Transferencias Hospitalarias por Dpto/Servicios"
                    Dim oRpt26 As New SIGHReportes.RptHTransferencia
                    oRpt26.EjecutaFormulario
                    Set oRpt26 = Nothing
             End Select
        Case lcXEMERGENCIA
             Select Case mrs_Reportes!reporte
             Case "Egresos Emergencia"
                    Dim oRptHosp1 As New SIGHProxies.clReportesEgreHosp
                    oRptHosp1.IdTipoReporte = sighentidades.sghReporteEgresosHospitalario
                    oRptHosp1.idTipoServicio = 1
                    oRptHosp1.EjecutaFormulario
                    Set oRptHosp1 = Nothing
             Case "Ingresos Emergencia"
                    Dim oRptIngHosp1 As New SIGHProxies.clReporteIngrHosp
                    oRptIngHosp1.IdTipoReporte = sighentidades.sghReporteIngresosHospitalario
                    oRptIngHosp1.idTipoServicio = 1
                    oRptIngHosp1.EjecutaFormulario
                    Set oRptIngHosp1 = Nothing
             Case "Primeras causas de Morbilidad por Emergencia, según ciclos de vida"
                    Dim oRpt225 As New SIGHReportes.RptHMorbEm
                    oRpt225.EjecutaFormulario
                    Set oRpt225 = Nothing
             End Select
        Case lcXLABORATORIO
             Select Case mrs_Reportes!reporte
             Case "Producción de Laboratorio"
                    Dim OrlRepProducPagoDeuda As New SIGHProxies.rlRepProducPagoDeuda
                    OrlRepProducPagoDeuda.idUsuario = ml_IdUsuarioAuditoria
                    OrlRepProducPagoDeuda.EjecutaFormulario
                    Set OrlRepProducPagoDeuda = Nothing
             Case "Productividad Consolidada"
                    Dim OrlRepProducPagoDeuda1 As New rlRepProducPagoDeuda1
                    OrlRepProducPagoDeuda1.idUsuario = ml_IdUsuarioAuditoria
                    OrlRepProducPagoDeuda1.EjecutaFormulario
                    Set OrlRepProducPagoDeuda1 = Nothing
             Case "Productividad de Laboratorio"
                    Dim OrlRepProduccion As New SIGHProxies.rlRepProduccion
                    OrlRepProduccion.idUsuario = ml_IdUsuarioAuditoria
                    OrlRepProduccion.EjecutaFormulario
                    Set OrlRepProduccion = Nothing
             Case "Pruebas Registradas"
                    Dim OrLabPruebas As New rLabPruebas
                    OrLabPruebas.idUsuario = ml_IdUsuarioAuditoria
                    OrLabPruebas.EjecutaFormulario
                    Set OrLabPruebas = Nothing
             Case "Pruebas Registradas con Resultados por Grupos"
                    Dim ORrlRepTipoAnalisisConRes As New rlRepTipoAnalisisConRes
                    ORrlRepTipoAnalisisConRes.idUsuario = ml_IdUsuarioAuditoria
                    ORrlRepTipoAnalisisConRes.EjecutaFormulario
                    Set ORrlRepTipoAnalisisConRes = Nothing
             Case "Pruebas Registradas por Grupos"
                    Dim ORrlRepTipoAnalisis As New SIGHProxies.rlRepTipoAnalisis
                    ORrlRepTipoAnalisis.idUsuario = ml_IdUsuarioAuditoria
                    ORrlRepTipoAnalisis.EjecutaFormulario
                    Set ORrlRepTipoAnalisis = Nothing
             End Select
        Case lcXIMAGENOLOGIA
             Select Case mrs_Reportes!reporte
             Case "Consumo de Insumos por Servicio"
                    Dim oRepConsumodeInsumosporServicios As New SIGHImagen.RepInsumoPorServicio
                    oRepConsumodeInsumosporServicios.idUsuario = ml_IdUsuarioAuditoria
                    oRepConsumodeInsumosporServicios.EjecutaFormulario
                    Set oRepConsumodeInsumosporServicios = Nothing
             Case "Consumo de Insumos por Tipo Servicio"
                    Dim oRepConsumodeInsumos As New SIGHImagen.RepInsumoPorTipoServ
                    oRepConsumodeInsumos.idUsuario = ml_IdUsuarioAuditoria
                    oRepConsumodeInsumos.EjecutaFormulario
                    Set oRepConsumodeInsumos = Nothing
             Case "Ecografía General por Fechas"
                    Dim oRepEcogGen As New SIGHImagen.RepEcogGen
                    oRepEcogGen.idUsuario = ml_IdUsuarioAuditoria
                    oRepEcogGen.EjecutaFormulario
                    Set oRepEcogGen = Nothing
             Case "Ecografía Obstétrica por Fechas"
                    Dim oRepEcogObs As New SIGHImagen.RepEcogObs
                    oRepEcogObs.idUsuario = ml_IdUsuarioAuditoria
                    oRepEcogObs.EjecutaFormulario
                    Set oRepEcogObs = Nothing
             Case "Kardex"
                    Dim oRepImgKardex As New SIGHImagen.RepKardex
                    oRepImgKardex.idUsuario = ml_IdUsuarioAuditoria
                    oRepImgKardex.EjecutaFormulario
                    Set oRepImgKardex = Nothing
             Case "Movimiento diario de Entradas y Salidas"
                    Dim oRepImgMovDiario As New SIGHImagen.RepMovimientoDiario
                    oRepImgMovDiario.idUsuario = ml_IdUsuarioAuditoria
                    oRepImgMovDiario.EjecutaFormulario
                    Set oRepImgMovDiario = Nothing
             Case "Producción, Pagos y Deudas"
                    Dim oRepProducciónPagosyDeuda As New SIGHImagen.RepProducPagoDeuda
                    oRepProducciónPagosyDeuda.idUsuario = ml_IdUsuarioAuditoria
                    oRepProducciónPagosyDeuda.EjecutaFormulario
                    Set oRepProducciónPagosyDeuda = Nothing
             Case "Productividad por Fechas"
                    Dim oRepProduccion As New SIGHImagen.RepProduccion
                    oRepProduccion.idUsuario = ml_IdUsuarioAuditoria
                    oRepProduccion.EjecutaFormulario
                    Set oRepProduccion = Nothing
             Case "Rayos X por Fechas"
                    Dim oRepRayosX As New SIGHImagen.RepRayosX
                    oRepRayosX.idUsuario = ml_IdUsuarioAuditoria
                    oRepRayosX.EjecutaFormulario
                    Set oRepRayosX = Nothing
             Case "Tomografía por Fechas"
                    Dim oRepTomografia As New SIGHImagen.RepTomografia
                    oRepTomografia.idUsuario = ml_IdUsuarioAuditoria
                    oRepTomografia.EjecutaFormulario
                    Set oRepTomografia = Nothing
             End Select
        Case lcXECONOMIA
             Select Case mrs_Reportes!reporte
             Case "Atenciones SIS en Hosp/Emerg/CE"
                    Dim oRptclAtencionesTotales As New SIGHProxies.clAtencionesTotales
                    oRptclAtencionesTotales.EjecutaFormulario
                    Set oRptclAtencionesTotales = Nothing
             Case "Cobertura y proceso de prestaciones de Salud del SIS"
                    Dim oRepSIS As New SIGHProxies.RptEconRepSIS
                    oRepSIS.idUsuario = ml_IdUsuarioAuditoria
                    oRepSIS.EjecutaFormulario
                    Set oRepSIS = Nothing
             Case "Consolidado de Ventas"
                    Dim RepConsolidadoVentas As New RpRegistroVentas
                    RepConsolidadoVentas.IdTipoReporte = 2
                    RepConsolidadoVentas.idUsuario = ml_IdUsuarioAuditoria
                    RepConsolidadoVentas.Show 1
                    Set RepConsolidadoVentas = Nothing
             Case "Consumo por Puntos de Carga"
                    Dim oRptConsPtoCarga As New SIGHReportes.RptEConsumoXptoCarga
                    oRptConsPtoCarga.idUsuario = ml_IdUsuarioAuditoria
                    oRptConsPtoCarga.EjecutaFormulario
                    Set oRptConsPtoCarga = Nothing
             Case "Consumo x Convenio"
                    Dim oRepConvenios As New rptEconRepConvenios
                    oRepConvenios.idUsuario = ml_IdUsuarioAuditoria
                    oRepConvenios.EjecutaFormulario
                    Set oRepConvenios = Nothing
             Case "Cuentas para Liquidación"
                    Dim oRptLiq As New SIGHReportes.RptESisSoatExoConv
                    oRptLiq.idUsuario = ml_IdUsuarioAuditoria
                    oRptLiq.EjecutaFormulario
                    Set oRptLiq = Nothing
             Case "Detalle por Partida"
                    Dim oRptPartidaDetalle As New RptEpartidaDetalle
                    oRptPartidaDetalle.EjecutaFormulario
                    Set oRptPartidaDetalle = Nothing
             Case "Exoneraciones General"
                    Dim oRpt229 As New SIGHReportes.RptEExoneraciones
                    oRpt229.EjecutaFormulario
                    Set oRpt229 = Nothing
             Case "Exoneraciones por Servicio, Especialiad, Dpto"
                    Dim oRpt239 As New SIGHReportes.RptEExoGeneral
                    oRpt239.EjecutaFormulario
                    Set oRpt239 = Nothing
             Case "Reembolsos Anuales"
                    Dim oRptERembolsoAnual As New RptERembolsoAnual
                    oRptERembolsoAnual.idUsuario = ml_IdUsuarioAuditoria
                    oRptERembolsoAnual.EjecutaFormulario
                    Set oRptERembolsoAnual = Nothing
             Case "Resumen por Partida"
                    Dim oRptResumenPartida As New RptEPartidaResumen
                    oRptResumenPartida.EjecutaFormulario
                    Set oRptResumenPartida = Nothing
             Case "Tipo Tarifa (CAJA)"
                    Dim oRptEtipoTarifa As New SIGHReportes.RptEtipoTarifa
                    oRptEtipoTarifa.EjecutaFormulario
                    Set oRptEtipoTarifa = Nothing
             End Select
        
        Case lcXFARMACIA
             Select Case mrs_Reportes!reporte
             Case "Saldos por Almacen"
                  Dim oRptFSaldos As New SighFarmacia.RepSaldosPorAlmacen
                  oRptFSaldos.idUsuario = ml_IdUsuarioAuditoria
                  oRptFSaldos.EjecutaFormulario
                  Set oRptFSaldos = Nothing
             Case "Kardex"
                  Dim oRptVtas As New SighFarmacia.RepKardex
                  oRptVtas.idUsuario = ml_IdUsuarioAuditoria
                  oRptVtas.EjecutaFormulario
                  Set oRptVtas = Nothing
             Case "Formato ICI"
                  Dim oRptICI As New SIGHProxies.RepICI
                  oRptICI.idUsuario = ml_IdUsuarioAuditoria
                  oRptICI.EjecutaFormulario
                  Set oRptICI = Nothing
             Case "Formato IDI"
                  Dim oRptIDI As New SighFarmacia.RepIDI
                  oRptIDI.idUsuario = ml_IdUsuarioAuditoria
                  oRptIDI.EjecutaFormulario
                  Set oRptIDI = Nothing
             Case "Movimiento de Documentos E/S"
                  Dim oRptMovES As New SighFarmacia.RepMovimientoES
                  oRptMovES.idUsuario = ml_IdUsuarioAuditoria
                  oRptMovES.EjecutaFormulario
                  Set oRptMovES = Nothing
             Case "Productos por Vencer"
                  Dim oRptProdXvencer As New SighFarmacia.RepProductoPorVencer
                  oRptProdXvencer.EjecutaFormulario
                  Set oRptProdXvencer = Nothing
             Case "Montos según Plan"
                  Dim oMontosP As New SighFarmacia.RepMontosXplan
                  oMontosP.idUsuario = ml_IdUsuarioAuditoria
                  oMontosP.EjecutaFormulario
                  Set oMontosP = Nothing
             Case "Recetas por Servicio"
                  Dim oRecetas As New SighFarmacia.RepRecetasXservicio
                  oRecetas.idUsuario = ml_IdUsuarioAuditoria
                  oRecetas.EjecutaFormulario
                  Set oRecetas = Nothing
             Case "Consumo por Cuenta"
                  Dim oConsCta As New SighFarmacia.RepConsumoPorCuenta
                  oConsCta.EjecutaFormulario
                  Set oConsCta = Nothing
             Case "Consumo promedio Anual"
                  Dim oConsAnual As New SighFarmacia.RepConsumoPromAnual
                  oConsAnual.idUsuario = ml_IdUsuarioAuditoria
                  oConsAnual.EjecutaFormulario
                  Set oConsAnual = Nothing
             Case "Recetas registradas por Usuario del Sistema"
                  Dim oRepXusuario As New SighFarmacia.RepRecetasXusuario
                  oRepXusuario.idUsuario = ml_IdUsuarioAuditoria
                  oRepXusuario.EjecutaFormulario
                  Set oRepXusuario = Nothing
             Case "Consumo por Servicios"
                  Dim oRepConsumoXservicio As New RepConsumoXservicio
                  oRepConsumoXservicio.EjecutaFormulario
                  Set oRepConsumoXservicio = Nothing
             Case "Ventas por productos según forma de pago"
                  Dim oRptFKardex As New SighFarmacia.RepMovimientoES
                  oRptFKardex.idUsuario = sighentidades.Usuario
                  oRptFKardex.EjecutaFrm
                  Set oRptFKardex = Nothing
             End Select
        Case lcXHERRAMIENTAS
             Select Case mrs_Reportes!reporte
             Case "Actualiza parametros"
                    Dim oActParametros As New HerrActualizacionParametros
                    oActParametros.Show 1
                    Set oActParametros = Nothing
             Case "Exporta/Importa datos a Citas Web"
                    Dim oHerrExportaCitasWeb As New HerrExportaCitasWeb
                    oHerrExportaCitasWeb.idUsuario = ml_IdUsuarioAuditoria
                    oHerrExportaCitasWeb.Show 1
                    Set oHerrExportaCitasWeb = Nothing
             Case "Exportar datos a SUNASA"
                    Dim oSUNASA As New SIGHProxies.clSunasa
                    oSUNASA.idUsuario = ml_IdUsuarioAuditoria
                    oSUNASA.lcNombrePc = lc_NombrePc
                    oSUNASA.MostrarFormulario
                    Set oSUNASA = Nothing
             Case "Pasa Atencion de NN hacia Paciente con Historia Clínica"
                    Dim oHerrModificaNN As New HerrModificaPacienteAtencionHE
                    oHerrModificaNN.idUsuario = ml_IdUsuarioAuditoria
                    oHerrModificaNN.Show 1
                    Set oHerrModificaNN = Nothing
             Case "Regenerar Saldos"
                    Dim oRegeneraSaldos As New SIGHProxies.RegeneraSaldos
                    oRegeneraSaldos.idUsuario = ml_IdUsuarioAuditoria
                    oRegeneraSaldos.lcNombrePc = lc_NombrePc
                    oRegeneraSaldos.MostrarFormulario
                    Set oRegeneraSaldos = Nothing
             Case "Reporte de Registro de Información por Usuario del Sistema"
                    Dim oRpt233 As New SIGHReportes.RptHerrUsuarioSistema
                    oRpt233.EjecutaFormulario
                    Set oRpt233 = Nothing
             Case "Reprogramación Médica"
                    Dim oHerrModifPac As New SIGHProxies.clHerrReprogramMedica
                    oHerrModifPac.idUsuario = ml_IdUsuarioAuditoria
                    oHerrModifPac.MostrarFormulario
                    Set oHerrModifPac = Nothing
             End Select
        End Select
    End If
    mrs_Reportes.Filter = ""
End Sub

Private Sub PVExplorer1_ItemClick(ByVal Item As Object)
    On Error GoTo errPV
    Dim lnFila As Integer
    For lnFila = 0 To 100
        If PVExplorer1.TreeView.SelectedNode.ListItems.Item(lnFila).Selected = True Then
           Dim oPrincipalc2 As New PrincipalC2
           oPrincipalc2.MuestraOpcionElegida = PVExplorer1.TreeView.SelectedNode.ListItems.Item(lnFila).SubItems(1)
           oPrincipalc2.ListBarItemClave = PVExplorer1.TreeView.SelectedNode.ListItems.Item(lnFila).SubItems(1)
           oPrincipalc2.Show 1
           Set oPrincipalc2 = Nothing
           Exit For
        End If
    Next
errPV:

End Sub


