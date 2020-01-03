VERSION 5.00
Begin VB.Form BienesInsumosBusqueda 
   Caption         =   "Busqueda de Productos"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "BienesInsumosBusqueda.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin Galenhos.ucBienesInsumosListaBus ucBienesInsumosListaBus1 
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9551
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   5385
      Width           =   9990
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "BienesInsumosBusqueda.frx":0CCA
         DownPicture     =   "BienesInsumosBusqueda.frx":112A
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
         Left            =   3495
         Picture         =   "BienesInsumosBusqueda.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BienesInsumosBusqueda.frx":1A14
         DownPicture     =   "BienesInsumosBusqueda.frx":1ED8
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
         Left            =   5040
         Picture         =   "BienesInsumosBusqueda.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "BienesInsumosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdProducto As Long
Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucBienesInsumosListaBus1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucBienesInsumosListaBus1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucBienesInsumosListaBus1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucBienesInsumosListaBus1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let IdTipoServicio(lValue As Long)
    ml_IdProducto = lValue
End Property
Property Get IdTipoServicio() As Long
    IdTipoServicio = ml_IdProducto
End Property
Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
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
    Me.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()

    Me.ucBienesInsumosListaBus1.Inicializar
    
    Me.ucBienesInsumosListaBus1.Titulo = "Búsqueda de Bienes Insumos"
    Me.ucBienesInsumosListaBus1.ConfigurarTiposDeBienesEInsumos
    ucBienesInsumosListaBus1.IdDepartamentoHospital = ml_IdDepartamentoHospital
    If mb_EjecutarBusquedaOnLoad Then
        ucBienesInsumosListaBus1.RealizarBusqueda
    End If
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucBienesInsumosListaBus1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub




