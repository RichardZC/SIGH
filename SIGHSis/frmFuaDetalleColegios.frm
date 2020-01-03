VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form FuaDetalleColegios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Instituciones Educativas"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   Icon            =   "frmFuaDetalleColegios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBuscaColegios 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame8 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   4080
         Width           =   6015
         Begin VB.CommandButton btnCancelar 
            Caption         =   "Cancelar (ESC)"
            DisabledPicture =   "frmFuaDetalleColegios.frx":000C
            DownPicture     =   "frmFuaDetalleColegios.frx":04D0
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   3360
            Picture         =   "frmFuaDetalleColegios.frx":09BC
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Aceptar (F2)"
            DisabledPicture =   "frmFuaDetalleColegios.frx":0EA8
            DownPicture     =   "frmFuaDetalleColegios.frx":1308
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1680
            Picture         =   "frmFuaDetalleColegios.frx":177D
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin UltraGrid.SSUltraGrid grdColegios 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         MaxColScrollRegions=   50
         MaxRowScrollRegions=   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdColegios"
      End
      Begin VB.Label Label3 
         Caption         =   "CÓDIGO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label 
         Caption         =   "INSTITUCIÓN EDUCATIVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   2235
      End
   End
End
Attribute VB_Name = "FuaDetalleColegios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: MINSA - Oficina de Informatica y Telecomunicaciones
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busqueda de Instituciones Educativas para el FUA v2
'        Programado por: Cachay F
'        Fecha: Agosto 2015
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mr_ReglasSIS As New SIGHSis.ReglasSISgalenhos   'Representa la Capa de Negocios del Modulo LAB GalenHos
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ms_CodigoColegio As String                              'Representa el codigo devuelto por el LAB LAB
Dim ms_DescColegio As String                          'Repsenta la descripcion del Codigo del LAB
Dim oRcs_DetalleColegio As New Recordset                 'Representa el detalle de codigo LAB de la base de datos
Dim ms_FiltroLAB As String                              'Representa el filtro de codigo LAB
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_Teclado As New sighentidades.Teclado

Property Let CodigoColegio(sValue As String)
   ms_CodigoColegio = sValue
End Property
Property Get CodigoColegio() As String
   CodigoColegio = ms_CodigoColegio
End Property

Property Let DescColegio(sValue As String)
   ms_DescColegio = sValue
End Property
Property Get DescColegio() As String
   DescColegio = ms_DescColegio
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let FiltroLAB(sValue As String)
   ms_FiltroLAB = sValue
End Property

Private Sub btnAgregar_Click()
    fraBuscaColegios.Enabled = False
End Sub

Private Sub btnCancelaAgregarColegio_Click()
    fraBuscaColegios.Enabled = True
    BuscarResultadoColegios
End Sub

Private Sub Form_Load()
    mo_Apariencia.ConfigurarFilasBiColores Me.grdColegios, sighentidades.GrillaConFilasBicolor
    txtCodigo.Text = ms_CodigoColegio
    Set oRcs_DetalleColegio = mr_ReglasSIS.SisFuaColegiosSeleccionarPorCodigoNombre(ms_CodigoColegio, "")
    Set grdColegios.DataSource = oRcs_DetalleColegio
End Sub

Private Sub txtDescripcio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.grdColegios
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.txtDescripcio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtDescripcio_KeyUp(KeyCode As Integer, Shift As Integer)
    BuscarResultadoColegios
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    BuscarResultadoColegios
End Sub

Sub BuscarResultadoColegios()
    mo_Apariencia.ConfigurarFilasBiColores Me.grdColegios, sighentidades.GrillaConFilasBicolor
    Set oRcs_DetalleColegio = mr_ReglasSIS.SisFuaColegiosSeleccionarPorCodigoNombre(Me.txtCodigo.Text, Me.txtDescripcio.Text)
    Set grdColegios.DataSource = oRcs_DetalleColegio
End Sub

Private Sub grdColegios_DblClick()
    btnAceptar_Click
End Sub

Private Sub grdColegios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
With Me.grdColegios.Bands(0)
    .Columns("CODIGO").Header.Caption = "Código"
    .Columns("CODIGO").Width = 1200
    .Columns("COLEGIO").Header.Caption = "Intitución Educativa"
    .Columns("COLEGIO").Width = 3200
    .Columns("UBIGEO").Header.Caption = "Ubigeo"
    .Columns("UBIGEO").Width = 2000
    .Columns("DIRECCION").Header.Caption = "Dirección"
    .Columns("DIRECCION").Width = 2500
End With
End Sub

'Actualizado 25092014
Private Sub grdColegios_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdColegios_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF3 Then
    btnAceptar_Click
End If
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    ms_CodigoColegio = CStr(Me.grdColegios.ActiveRow.Cells("CODIGO").Value)
    ms_DescColegio = CStr(Me.grdColegios.ActiveRow.Cells("COLEGIO").Value)
    Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ms_CodigoColegio = ""
    ms_DescColegio = ""
    Visible = False
End Sub

Public Sub MostrarFormulario()
    Me.Show 1
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

