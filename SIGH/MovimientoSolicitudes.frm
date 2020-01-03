VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form MovimientoSolicitudes 
   Caption         =   "Historias Solicitadas Adicionales"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   Icon            =   "MovimientoSolicitudes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   3
      Top             =   3480
      Width           =   10545
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MovimientoSolicitudes.frx":0CCA
         DownPicture     =   "MovimientoSolicitudes.frx":112A
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
         Left            =   3765
         Picture         =   "MovimientoSolicitudes.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MovimientoSolicitudes.frx":1A14
         DownPicture     =   "MovimientoSolicitudes.frx":1ED8
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
         Left            =   5310
         Picture         =   "MovimientoSolicitudes.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdPrestamosHC 
      Height          =   3000
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   5292
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de solicitud de historias"
   End
   Begin VB.Label Label1 
      Caption         =   "El paciente ingresado tiene las siguientes solicitudes adicionales, por favor indique cual es la solicitud que desea procesar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   10515
   End
End
Attribute VB_Name = "MovimientoSolicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Movimiento de Solicitud de Historia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_HistoriasSolicitadas As Recordset
Dim mb_Aceptar As Boolean
Dim mo_Apariencia As New GridInfragistic

Property Set HistoriasSolicitadas(rsHistorias As Recordset)
    Set mrs_HistoriasSolicitadas = rsHistorias
End Property
Property Get HistoriasSolicitadas() As Recordset
    Set HistoriasSolicitadas = mrs_HistoriasSolicitadas
End Property
Property Let Aceptar(bValue As Boolean)
    mb_Aceptar = bValue
End Property
Property Get Aceptar() As Boolean
    Aceptar = mb_Aceptar
End Property

Private Sub btnAceptar_Click()
    mb_Aceptar = True
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    mb_Aceptar = False
    Me.Visible = False
End Sub

Private Sub Form_Load()
        
        Set grdPrestamosHC.DataSource = mrs_HistoriasSolicitadas
        mo_Apariencia.ConfigurarFilasBiColores grdPrestamosHC, sighEntidades.GrillaConFilasBicolor

End Sub

Private Sub grdPrestamosHC_AfterRowActivate()
    Set mrs_HistoriasSolicitadas = grdPrestamosHC.DataSource
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub grdPrestamosHC_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    AdministrarKeyPreview KeyCode.Value
End Sub
