VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form AHCSinDevolver 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10170
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
   ScaleHeight     =   6420
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkExcel 
      Alignment       =   1  'Right Justify
      Caption         =   "En Excel"
      Height          =   315
      Left            =   90
      Picture         =   "AHCSinDevolver.frx":0000
      TabIndex        =   5
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
      TabIndex        =   2
      Top             =   5175
      Width           =   10035
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCSinDevolver.frx":0312
         DownPicture     =   "AHCSinDevolver.frx":07D6
         Height          =   705
         Left            =   4905
         Picture         =   "AHCSinDevolver.frx":0CC2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCSinDevolver.frx":11AE
         DownPicture     =   "AHCSinDevolver.frx":160E
         Height          =   705
         Left            =   3375
         Picture         =   "AHCSinDevolver.frx":1A83
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.CommandButton btnBuscar 
      Height          =   315
      Left            =   7005
      Picture         =   "AHCSinDevolver.frx":1EF8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   1305
   End
   Begin UltraGrid.SSUltraGrid grdHistoriasC 
      Height          =   4560
      Left            =   45
      TabIndex        =   0
      Top             =   570
      Width           =   9930
      _ExtentX        =   17515
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
      Caption         =   "Lista de historias clínicas"
   End
End
Attribute VB_Name = "AHCSinDevolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SCCQ 28/08/2020 Cambio27 Inicio
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
            'oRptClaseCry.Dias = "3"
            'oRptClaseCry.FechaInicio = "01/01/2020"
            oRptClaseCry.TextoDelFiltro = "Solicitadas por TRAMITES ADMINISTRATIVOS sin devolución mayor a 72 horas"
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
   Me.MousePointer = 1
End Sub

Private Sub btnBuscar_Click()
 Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub RealizarBusqueda()
Set grdHistoriasC.DataSource = mo_ReglasAC.SeleccionarHCSinDevolver(72)
End Sub
Private Sub Form_Load()
   If ml_mostrarReporte = True Then
       'btnBuscar_Click
        RealizarBusqueda
    End If
End Sub
'SCCQ 28/08/2020 Cambio27 Fin
