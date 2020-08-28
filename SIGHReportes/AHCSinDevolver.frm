VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form AHCSinDevolver 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
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
   ScaleHeight     =   5985
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnBuscar 
      Height          =   315
      Left            =   7005
      Picture         =   "AHCSinDevolver.frx":0000
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
Private Sub btnBuscar_Click()
 Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Set grdHistoriasC.DataSource = mo_ReglasAC.SeleccionarHCSinDevolver(72)
End Sub
'SCCQ 28/08/2020 Cambio27 Fin
