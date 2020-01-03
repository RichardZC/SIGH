VERSION 5.00
Begin VB.Form BuscarCodigosSunat 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7560
   ClientLeft      =   8940
   ClientTop       =   4170
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin SIGHNegocios.UcCodigosSunat UcCodigosSunat1 
      Height          =   6330
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   11165
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   10920
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BuscarCodigosSunat.frx":0000
         DownPicture     =   "BuscarCodigosSunat.frx":0460
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
         Left            =   4080
         Picture         =   "BuscarCodigosSunat.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BuscarCodigosSunat.frx":0D4A
         DownPicture     =   "BuscarCodigosSunat.frx":120E
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
         Left            =   5625
         Picture         =   "BuscarCodigosSunat.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "BuscarCodigosSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca procedimiento CPT
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdTipoCatalogo As Long
'Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean
Dim ml_IdPuntoCarga As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_TipoServicioOfrecido As Long
Dim ms_codigoSunat As String

Property Let codigoSunat(lValue As String)
    ms_codigoSunat = lValue
End Property
Property Get codigoSunat() As String
    codigoSunat = ms_codigoSunat
End Property



Property Set DataSource(oValue As ADODB.Recordset)
    Set UcCodigosSunat1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UcCodigosSunat1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    UcCodigosSunat1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = UcCodigosSunat1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property



Property Let EjecutarBusquedaOnLoad(bValue As Boolean)
    mb_EjecutarBusquedaOnLoad = bValue
End Property
Property Get EjecutarBusquedaOnLoad() As Boolean
    EjecutarBusquedaOnLoad = mb_EjecutarBusquedaOnLoad
End Property

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub
Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    IdRegistroSeleccionado = 0
    Me.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    UcCodigosSunat1.inicializar
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            UcCodigosSunat1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub UcCodigosSunat1_SeleccionaRegistro(lcCodigoSunat As String)
    If lcCodigoSunat <> "" Then
       ms_codigoSunat = lcCodigoSunat
       btnAceptar_Click
    End If
End Sub
