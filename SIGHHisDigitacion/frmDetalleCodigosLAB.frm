VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmDetalleCodigosLAB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de código LAB"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   Icon            =   "frmDetalleCodigosLAB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcioLab 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtValores 
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
      Left            =   0
      MaxLength       =   4
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   5895
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmDetalleCodigosLAB.frx":000C
         DownPicture     =   "frmDetalleCodigosLAB.frx":046C
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
         Left            =   1440
         Picture         =   "frmDetalleCodigosLAB.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleCodigosLAB.frx":0D56
         DownPicture     =   "frmDetalleCodigosLAB.frx":121A
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
         Left            =   3120
         Picture         =   "frmDetalleCodigosLAB.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleCodigoLAB 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
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
      Caption         =   "ugvDetalleCodigoLAB"
   End
   Begin VB.Label Label 
      Caption         =   "Nombre Codigo Lab"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "Código"
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
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   675
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Códigos LAB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmDetalleCodigosLAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica de Listado de Codigos LAB.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mr_ReglasLAB As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo LAB GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ms_CodigoLab As String                              'Representa el codigo devuelto por el LAB LAB
Dim ms_DescCodigoLAB As String                          'Repsenta la descripcion del Codigo del LAB
Dim oRcs_DetalleLAB As New Recordset                 'Representa el detalle de codigo LAB de la base de datos
Dim ms_FiltroLAB As String                              'Representa el filtro de codigo LAB
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_Teclado As New SIGHEntidades.Teclado

Property Let CodigoLab(sValue As String)
   ms_CodigoLab = sValue
End Property
Property Get CodigoLab() As String
   CodigoLab = ms_CodigoLab
End Property

Property Let DescripcionCodigoLAB(sValue As String)
   ms_DescCodigoLAB = sValue
End Property
Property Get DescripcionCodigoLAB() As String
   DescripcionCodigoLAB = ms_CodigoLab
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let FiltroLAB(sValue As String)
   ms_FiltroLAB = sValue
End Property
'MODIFICADO POR YEPE NOVIEMBRE
Private Sub Form_Load()
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoLAB, SIGHEntidades.GrillaConFilasBicolor
    'Set oRcs_DetalleLABLAB = mr_ReglasLAB.ObtenerListaCodigosLAB '("","")
    Set oRcs_DetalleLAB = mr_ReglasLAB.ObtenerListaCodigosLABporCodigoyNombre(ms_CodigoLab, "")
    Set ugvDetalleCodigoLAB.DataSource = oRcs_DetalleLAB
End Sub

Private Sub txtDescripcioLab_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.ugvDetalleCodigoLAB
    AdministrarKeyPreview KeyCode
End Sub

'MODIFICADO POR YEPE NOVIEMBRE
Private Sub txtDescripcioLab_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoLAB, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleLAB = mr_ReglasLAB.ObtenerListaCodigosLABporCodigoyNombre(Me.txtValores.Text, Me.txtDescripcioLab.Text)
    Set ugvDetalleCodigoLAB.DataSource = oRcs_DetalleLAB
End Sub

Private Sub txtValores_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.txtDescripcioLab
    AdministrarKeyPreview KeyCode
End Sub

'MODIFICADO POR YEPE NOVIEMBRE
Private Sub txtValores_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoLAB, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleLAB = mr_ReglasLAB.ObtenerListaCodigosLABporCodigoyNombre(Me.txtValores.Text, Me.txtDescripcioLab.Text)
    Set ugvDetalleCodigoLAB.DataSource = oRcs_DetalleLAB
End Sub

Private Sub ugvDetalleCodigoLAB_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleCodigoLAB_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
With Me.ugvDetalleCodigoLAB.Bands(0)
    .Columns("IdHisSituacio").Hidden = True
    .Columns("valores").Header.Caption = "Codigo"
    .Columns("valores").Width = 1500
    .Columns("descripcio").Header.Caption = "Descripcion"
    .Columns("descripcio").Width = 2000
    .Columns("codigo").Hidden = True
    .Columns("est").Hidden = True
End With
End Sub

Private Sub ugvDetalleCodigoLAB_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF3 Then
    btnAceptar_Click
End If
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    ms_CodigoLab = CStr(Me.ugvDetalleCodigoLAB.ActiveRow.Cells("valores").Value)
    DescripcionCodigoLAB = CStr(Me.ugvDetalleCodigoLAB.ActiveRow.Cells("descripcio").Value)
    Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ms_CodigoLab = ""
    DescripcionCodigoLAB = ""
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

