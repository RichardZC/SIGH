VERSION 5.00
Begin VB.Form CatalogoBienesInsumosBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin GalenHos.ucCatBienesInsumosListaBus ucCatalogoBienesInsumosListaBus1 
      Height          =   7215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12726
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   7320
      Width           =   10935
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoBienesInsumosBusqueda.frx":0000
         DownPicture     =   "CatalogoBienesInsumosBusqueda.frx":04C4
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
         Picture         =   "CatalogoBienesInsumosBusqueda.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoBienesInsumosBusqueda.frx":0E9C
         DownPicture     =   "CatalogoBienesInsumosBusqueda.frx":12FC
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
         Picture         =   "CatalogoBienesInsumosBusqueda.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoBienesInsumosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdTipoCatalogo As Long
'Dim ml_IdDepartamentoHospital As Long
Dim mb_EjecutarBusquedaOnLoad As Boolean

Property Set DataSource(oValue As ADODB.Recordset)
    Set Me.ucCatalogoBienesInsumosListaBus1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucCatalogoBienesInsumosListaBus1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucCatalogoBienesInsumosListaBus1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucCatalogoBienesInsumosListaBus1.IdRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let IdTipoCatalogo(lValue As Long)
    ml_IdTipoCatalogo = lValue
End Property
Property Get IdTipoCatalogo() As Long
    IdTipoCatalogo = ml_IdTipoCatalogo
End Property
Property Let HabilitarTipoCatalogo(lValue As Boolean)
    ucCatalogoBienesInsumosListaBus1.HabilitarTipoCatalogo = lValue
End Property
Property Get HabilitarTipoCatalogo() As Boolean
    HabilitarTipoCatalogo = ucCatalogoBienesInsumosListaBus1.HabilitarTipoCatalogo
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
    
    
    ucCatalogoBienesInsumosListaBus1.Inicializar
    Me.ucCatalogoBienesInsumosListaBus1.Titulo = "Búsqueda de Catalogo de Bienes e Insumos"
    Me.ucCatalogoBienesInsumosListaBus1.ConfigurarTiposCatalogos
    ucCatalogoBienesInsumosListaBus1.IdTipoCatalogo = ml_IdTipoCatalogo
    ucCatalogoBienesInsumosListaBus1.SeleccionaTipoCatalogo
    'ucCatalogoBienesInsumosListaBus1.iºº.IdDepartamentoHospital = ml_IdDepartamentoHospital
    If mb_EjecutarBusquedaOnLoad Then
        ucCatalogoBienesInsumosListaBus1.RealizarBusqueda
    End If
    
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucCatalogoBienesInsumosListaBus1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub






