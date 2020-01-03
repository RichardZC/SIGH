VERSION 5.00
Begin VB.Form CamasBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de camas"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   Icon            =   "CamasBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SISGalenPlus.ucCamasLista ucCamasLista1 
      Height          =   5655
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9975
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Top             =   5670
      Width           =   10485
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CamasBusqueda.frx":0CCA
         DownPicture     =   "CamasBusqueda.frx":118E
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
         Left            =   4470
         Picture         =   "CamasBusqueda.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   255
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CamasBusqueda.frx":1B66
         DownPicture     =   "CamasBusqueda.frx":1FC6
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
         Left            =   2925
         Picture         =   "CamasBusqueda.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   255
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CamasBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Camas
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_idServicio As Long
Dim ml_IdTipoServicio As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucCamasLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucCamasLista1.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ucCamasLista1.idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ucCamasLista1.idRegistroSeleccionado
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let idTipoServicio(lValue As Long)
    ml_IdTipoServicio = lValue
End Property
Property Get idTipoServicio() As Long
    idTipoServicio = ml_IdTipoServicio
End Property
Property Let IdServicio(lValue As Long)
    ml_idServicio = lValue
End Property
Property Get IdServicio() As Long
    IdServicio = ml_idServicio
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

    Me.ucCamasLista1.inicializar
    Me.ucCamasLista1.Titulo = "Disponibilidad de camas"
    Me.ucCamasLista1.ConfigurarTipoServicio
    
    ucCamasLista1.idTipoServicio = ml_IdTipoServicio
    ucCamasLista1.IdServicio = ml_idServicio
    ucCamasLista1.ClicEnBotonBuscar
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucCamasLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub ucCamasLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
     If lnIdRegistroSeleccionado > 0 Then
        btnAceptar_Click
     End If
End Sub
