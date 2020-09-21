VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form AHCSinDevolver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HC sin devolución por trámites administrativos pasadas las 72 horas"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExcel 
      Alignment       =   1  'Right Justify
      Caption         =   "En Excel"
      Height          =   315
      Left            =   90
      Picture         =   "AHCSinDevolver.frx":0000
      TabIndex        =   4
      Top             =   75
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   75
      TabIndex        =   1
      Top             =   5655
      Width           =   14820
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCSinDevolver.frx":0312
         DownPicture     =   "AHCSinDevolver.frx":07D6
         Height          =   705
         Left            =   7530
         Picture         =   "AHCSinDevolver.frx":0CC2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCSinDevolver.frx":11AE
         DownPicture     =   "AHCSinDevolver.frx":160E
         Height          =   705
         Left            =   6000
         Picture         =   "AHCSinDevolver.frx":1A83
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdHistoriasC 
      Height          =   4560
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   14790
      _ExtentX        =   26088
      _ExtentY        =   8043
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de historias clínicas sin devolución por trámites administrativos pasadas las 72 horas"
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   6
      Top             =   5280
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   480
   End
End
Attribute VB_Name = "AHCSinDevolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SCCQ 08/09/2020 Cambio27 Inicio
Option Explicit
Dim mo_ReglasAC As New SIGHNegocios.ReglasArchivoClinico
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_mostrarReporte As Boolean
Property Let mostrarReporte(lValue As Boolean)
    ml_mostrarReporte = lValue
End Property
Private Sub btnAceptar_Click()
   Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.TextoDelFiltro = "Solicitadas por TRAMITES ADMINISTRATIVOS sin devolución mayor a 72 horas"
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
   Me.MousePointer = 1
End Sub
Public Sub RealizarBusqueda()
'Set grdHistoriasC.DataSource = mo_ReglasAC.SeleccionarHCSinDevolver(72)
Dim otabla As Recordset
Set otabla = mo_ReglasAC.SeleccionarHCSinDevolver(72)
Set grdHistoriasC.DataSource = otabla
lblTotal = otabla.RecordCount()
Set otabla = Nothing
'Data1.Recordset.RecordCount
End Sub
Private Sub btnCancelar_Click()
 Me.Visible = False
End Sub
Private Sub Form_Initialize()
mo_Apariencia.ConfigurarFilasBiColores grdHistoriasC, sighEntidades.GrillaConFilasBicolor
End Sub
Private Sub Form_Load()
RealizarBusqueda
End Sub
Private Sub grdHistoriasC_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
grdHistoriasC.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
grdHistoriasC.Bands(0).Columns("NroHistoriaClinica").Width = 1300

grdHistoriasC.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
grdHistoriasC.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
grdHistoriasC.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
grdHistoriasC.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
grdHistoriasC.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
grdHistoriasC.Bands(0).Columns("PrimerNombre").Width = 1500

grdHistoriasC.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
grdHistoriasC.Bands(0).Columns("SegundoNombre").Width = 1500

grdHistoriasC.Bands(0).Columns("fecha_prestada").Header.Caption = "Fecha préstamo"
grdHistoriasC.Bands(0).Columns("fecha_prestada").Width = 1900

grdHistoriasC.Bands(0).Columns("destino").Header.Caption = "Destino"
grdHistoriasC.Bands(0).Columns("destino").Width = 2500

grdHistoriasC.Bands(0).Columns("Observacion").Header.Caption = "Observación"
grdHistoriasC.Bands(0).Columns("Observacion").Width = 2500
End Sub
'SCCQ 08/09/2020 Cambio27 Fin
