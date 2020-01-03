VERSION 5.00
Begin VB.Form ListaSaldos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ListaSaldos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SighFarmacia.ucListaSaldos ucListaSaldos1 
      Height          =   5325
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   13515
      _ExtentX        =   17171
      _ExtentY        =   9393
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   5505
      Width           =   13470
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ListaSaldos.frx":0CCA
         DownPicture     =   "ListaSaldos.frx":112A
         Height          =   700
         Left            =   5377
         Picture         =   "ListaSaldos.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ListaSaldos.frx":1A14
         DownPicture     =   "ListaSaldos.frx":1ED8
         Height          =   700
         Left            =   6922
         Picture         =   "ListaSaldos.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ListaSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Lista Saldos
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Property Let CantidadSeleccionado(lValue As Long)
    ucListaSaldos1.CantidadSeleccionado = lValue
End Property
Property Get CantidadSeleccionado() As Long
    CantidadSeleccionado = ucListaSaldos1.CantidadSeleccionado
End Property
Property Let CodigoSeleccionado(lValue As String)
    ucListaSaldos1.CodigoSeleccionado = lValue
End Property
Property Get CodigoSeleccionado() As String
    CodigoSeleccionado = ucListaSaldos1.CodigoSeleccionado
End Property
Property Let NombreSeleccionado(lValue As String)
    ucListaSaldos1.NombreSeleccionado = lValue
End Property
Property Get NombreSeleccionado() As String
    NombreSeleccionado = ucListaSaldos1.NombreSeleccionado
End Property
Property Set DataSource(oValue As ADODB.Recordset)
    Set ucListaSaldos1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucListaSaldos1.DataSource
End Property

Property Let IdProductoSeleccionado(lValue As Long)
    ucListaSaldos1.IdProductoSeleccionado = lValue
End Property
Property Get IdProductoSeleccionado() As Long
    IdProductoSeleccionado = ucListaSaldos1.IdProductoSeleccionado
End Property
Property Let IdAlmacenSeleccionado(lValue As Long)
    ucListaSaldos1.IdAlmacenSeleccionado = lValue
End Property
Property Get IdAlmacenSeleccionado() As Long
    IdAlmacenSeleccionado = ucListaSaldos1.IdAlmacenSeleccionado
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
    Me.ucListaSaldos1.Titulo = "Búsqueda de Producto"
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucListaSaldos1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Private Sub ucListaSaldos1_OnClick(IdAlmacenSeleccionado As Long, IdProductoSeleccionado As Long, lcCodigoSeleccionado As String, lcNombreSeleccionado As String, lnCantidad As Long)
    If IdAlmacenSeleccionado > 0 And IdProductoSeleccionado > 0 Then
       btnAceptar_Click
    End If
End Sub


