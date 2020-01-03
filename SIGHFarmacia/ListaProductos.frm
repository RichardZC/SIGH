VERSION 5.00
Begin VB.Form ListaProductos 
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
   Icon            =   "ListaProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SighFarmacia.ucListaProductos ucListaProductos1 
      Height          =   5460
      Left            =   45
      TabIndex        =   3
      Top             =   15
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   9631
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
      Left            =   105
      TabIndex        =   0
      Top             =   5520
      Width           =   13470
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ListaProductos.frx":0CCA
         DownPicture     =   "ListaProductos.frx":112A
         Height          =   700
         Left            =   5377
         Picture         =   "ListaProductos.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ListaProductos.frx":1A14
         DownPicture     =   "ListaProductos.frx":1ED8
         Height          =   700
         Left            =   6922
         Picture         =   "ListaProductos.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "ListaProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca productos
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lb_MuestraTodosItems As Boolean
Dim mi_BotonPresionado As sghBotonDetallePresionado
Property Let MuestraTodosItems(lValue As Boolean)
    lb_MuestraTodosItems = lValue
    Me.ucListaProductos1.MuestraTodosItems = lb_MuestraTodosItems
End Property

Property Let CodigoSeleccionado(lValue As String)
    ucListaProductos1.CodigoSeleccionado = lValue
End Property
Property Get CodigoSeleccionado() As String
    CodigoSeleccionado = ucListaProductos1.CodigoSeleccionado
End Property
Property Let NombreSeleccionado(lValue As String)
    ucListaProductos1.NombreSeleccionado = lValue
End Property
Property Get NombreSeleccionado() As String
    NombreSeleccionado = ucListaProductos1.NombreSeleccionado
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucListaProductos1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucListaProductos1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucListaProductos1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucListaProductos1.IdRegistroSeleccionado
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
    Me.ucListaProductos1.Titulo = "Búsqueda de Producto"
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub ucListaProductos1_OnClick(IdSeleccionado As Long, lcCodigoSeleccionado As String, lcNombreSeleccionado As String)
    If IdSeleccionado > 0 Then
       btnAceptar_Click
    End If
End Sub


