VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmDetalleCodigoPais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de código de nacionalidad"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   Icon            =   "frmDetalleCodigoPais.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigoPais 
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
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtNombrePais 
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
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   6015
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleCodigoPais.frx":000C
         DownPicture     =   "frmDetalleCodigoPais.frx":04D0
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
         Picture         =   "frmDetalleCodigoPais.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmDetalleCodigoPais.frx":0EA8
         DownPicture     =   "frmDetalleCodigoPais.frx":1308
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
         Picture         =   "frmDetalleCodigoPais.frx":177D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleCodigoNac 
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5106
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
      Caption         =   "ugvDetalleCodigoNac"
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
      TabIndex        =   8
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label 
      Caption         =   "Nombre País"
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
      TabIndex        =   7
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Códigos de Nacionalidad"
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
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmDetalleCodigoPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica de Listado de Codigos de Nacionalidad.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_IdPais As Long
Dim ms_CodigoNac As String                              'Representa el codigo devuelto por el HIS LAB
Dim ms_NombrePais As String                             'Repsenta la descripcion del Codigo del HIS
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim oRcs_DetalleCodNac As New Recordset
Dim mo_Teclado As New SIGHEntidades.Teclado

Property Let IdPais(lValue As Long)
   ml_IdPais = lValue
End Property
Property Get IdPais() As Long
   IdPais = ml_IdPais
End Property

Property Let CodigoNac(sValue As String)
   ms_CodigoNac = sValue
End Property
Property Get CodigoNac() As String
   CodigoNac = ms_CodigoNac
End Property

Property Let NombrePais(sValue As String)
   ms_NombrePais = sValue
End Property
Property Get NombrePais() As String
   NombrePais = ms_NombrePais
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub Form_Load()
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoNac, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleCodNac = mr_ReglasHIS.ObtenerListaCodigosNaciones("", "")
    Set ugvDetalleCodigoNac.DataSource = oRcs_DetalleCodNac
End Sub

Private Sub txtCodigoPais_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.txtNombrePais
    AdministrarKeyPreview KeyCode
End Sub

'Actuallizado Yamill 28102013 Inicio
Private Sub txtCodigoPais_KeyUp(KeyCode As Integer, Shift As Integer)
        mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoNac, SIGHEntidades.GrillaConFilasBicolor
        Set oRcs_DetalleCodNac = mr_ReglasHIS.ObtenerListaCodigosNaciones(txtCodigoPais.Text, txtNombrePais.Text)
        Set ugvDetalleCodigoNac.DataSource = oRcs_DetalleCodNac
End Sub

Private Sub txtNombrePais_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.ugvDetalleCodigoNac
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNombrePais_KeyUp(KeyCode As Integer, Shift As Integer)
        mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoNac, SIGHEntidades.GrillaConFilasBicolor
        Set oRcs_DetalleCodNac = mr_ReglasHIS.ObtenerListaCodigosNaciones(txtCodigoPais.Text, txtNombrePais.Text)
        Set ugvDetalleCodigoNac.DataSource = oRcs_DetalleCodNac
End Sub
'Actuallizado Yamill 28102013 Fin

Private Sub ugvDetalleCodigoNac_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleCodigoNac_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
With Me.ugvDetalleCodigoNac.Bands(0)
    .Columns("Idpais").Hidden = True
    .Columns("Codigo").Header.Caption = "Codigo"
    .Columns("Codigo").Width = 1000
    .Columns("nombre").Header.Caption = "Nombre pais"
    .Columns("nombre").Width = 2000
End With
End Sub

Private Sub ugvDetalleCodigoNac_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF3 Then
    btnAceptar_Click
Else
    AdministrarKeyPreview CInt(KeyCode)
End If
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    ml_IdPais = CInt(Me.ugvDetalleCodigoNac.ActiveRow.Cells("IdPais").Value)
    ms_CodigoNac = CStr(Me.ugvDetalleCodigoNac.ActiveRow.Cells("Codigo").Value)
    ms_NombrePais = CStr(Me.ugvDetalleCodigoNac.ActiveRow.Cells("nombre").Value)
    Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ml_IdPais = 0
    ms_CodigoNac = ""
    ms_NombrePais = ""
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

