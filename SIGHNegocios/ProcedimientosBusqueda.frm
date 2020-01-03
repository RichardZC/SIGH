VERSION 5.00
Begin VB.Form ProcedimientosBusqueda 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   2940
   ClientTop       =   1920
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   9195
   Begin SIGHNegocios.ucProcedimientosLista ucProcedimientosLista1 
      Height          =   4995
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   8811
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   45
      TabIndex        =   0
      Top             =   5085
      Width           =   9120
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ProcedimientosBusqueda.frx":0000
         DownPicture     =   "ProcedimientosBusqueda.frx":04C4
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
         Picture         =   "ProcedimientosBusqueda.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ProcedimientosBusqueda.frx":0E9C
         DownPicture     =   "ProcedimientosBusqueda.frx":12FC
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
         Left            =   3015
         Picture         =   "ProcedimientosBusqueda.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ProcedimientosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca procedimiento
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mi_IdDiferenciacion As Integer

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucProcedimientosLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucProcedimientosLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucProcedimientosLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucProcedimientosLista1.IdRegistroSeleccionado
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


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucProcedimientosLista1.Titulo = "Búsqueda de procedimientos"
    Me.ucProcedimientosLista1.IdDiferenciacion = mi_IdDiferenciacion
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucProcedimientosLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Public Property Get IdDiferenciacion() As Integer
    IdDiferenciacion = mi_IdDiferenciacion
End Property

Public Property Let IdDiferenciacion(ByVal iValue As Integer)
    mi_IdDiferenciacion = iValue
End Property
