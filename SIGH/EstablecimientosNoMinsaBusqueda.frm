VERSION 5.00
Begin VB.Form EstablecimientosNoMinsaBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11955
   Icon            =   "EstablecimientosNoMinsaBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GalenHos.ucEstablecimientosNoMinsaLista ucEstablecimientosNoMinsaLista1 
      Height          =   4725
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8334
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   60
      TabIndex        =   4
      Top             =   4710
      Width           =   11835
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         DisabledPicture =   "EstablecimientosNoMinsaBusqueda.frx":0CCA
         DownPicture     =   "EstablecimientosNoMinsaBusqueda.frx":10B3
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
         Left            =   180
         Picture         =   "EstablecimientosNoMinsaBusqueda.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Agregar"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EstablecimientosNoMinsaBusqueda.frx":18CB
         DownPicture     =   "EstablecimientosNoMinsaBusqueda.frx":1D8F
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
         Left            =   6105
         Picture         =   "EstablecimientosNoMinsaBusqueda.frx":227B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EstablecimientosNoMinsaBusqueda.frx":2767
         DownPicture     =   "EstablecimientosNoMinsaBusqueda.frx":2BC7
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
         Left            =   4560
         Picture         =   "EstablecimientosNoMinsaBusqueda.frx":303C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EstablecimientosNoMinsaBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_idUsuario As Long
Dim mo_lcNombrePc As String
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Set DataSource(oValue As ADODB.Recordset)
    Set ucEstablecimientosNoMinsaLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucEstablecimientosNoMinsaLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucEstablecimientosNoMinsaLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucEstablecimientosNoMinsaLista1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub

Private Sub cmdAgregar_Click()
        Dim mo_EstablecimientoNoMinsaDetalle As New EstablecimientoNoMinsaDetalle
        mo_EstablecimientoNoMinsaDetalle.Opcion = sghAgregar
        mo_EstablecimientoNoMinsaDetalle.idUsuario = ml_idUsuario
        mo_EstablecimientoNoMinsaDetalle.lnIdTablaLISTBARITEMS = 1204
        mo_EstablecimientoNoMinsaDetalle.lcNombrePc = mo_lcNombrePc
        mo_EstablecimientoNoMinsaDetalle.Show 1
        Unload mo_EstablecimientoNoMinsaDetalle

End Sub

Private Sub Form_Initialize()
    ucEstablecimientosNoMinsaLista1.ConfigurarEstablecimientos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    ucEstablecimientosNoMinsaLista1.Inicializar
    Me.ucEstablecimientosNoMinsaLista1.Titulo = "Búsqueda de Establecimientos no MINSA"
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucEstablecimientosNoMinsaLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

