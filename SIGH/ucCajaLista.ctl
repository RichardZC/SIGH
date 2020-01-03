VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucCajaLista 
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10080
   ScaleHeight     =   6255
   ScaleWidth      =   10080
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   75
      TabIndex        =   5
      Top             =   540
      Width           =   9930
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   3585
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
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1125
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   6555
         Picture         =   "ucCajaLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5190
         Picture         =   "ucCajaLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código           Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   210
         Width           =   6975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdCajas 
      Height          =   4350
      Left            =   120
      TabIndex        =   4
      Top             =   1815
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   7673
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de Cajas"
   End
   Begin VB.Label Label 
      Caption         =   "Obs: Debe configurar el formato de la SERIE de las FACTURAS (ejm: FF01, F001) y BOLETAS (ejm: BB01, B001)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   9855
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Listado de Cajas"
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
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   9975
   End
End
Attribute VB_Name = "ucCajaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar Cajas
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdCajas.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdCajas.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoFiltro(lValue As sghTipoFiltroPacientes)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroPacientes
    TipoFiltro = ml_TipoFiltro
End Property
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
    Dim oCaja As New DOCajaCaja
    oCaja.Codigo = txtCodigo.Text
    oCaja.Descripcion = txtDescripcion.Text

    Set grdCajas.DataSource = mo_AdminCaja.RealizarFiltroCajas(oCaja)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox mo_AdminCaja.MensajeError, vbInformation, "Filtro de cajas"
    End If
    'mo_Apariencia.ConfigurarFilasBiColores grdCajas, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtCodigo = ""
    UserControl.txtDescripcion = ""
End Sub
Private Sub cmbCajero_Click()
    'mo_cmbCajero.BoundColumn = "IdCajero"
    'mo_cmbCajero.ListField = "NombreCompleto"
    'Set mo_cmbCajero.RowSource = mo_AdminCaja.ServiciosSeleccionarPorTipoV2(Val(mo_cmbCajero.BoundText))
End Sub

Private Sub grdCajas_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCajas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCaja")
    

End Sub

Private Sub grdCajas_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCajas.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdCaja")
    
End Sub


Private Sub grdCajas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdCajas.Bands(0).Columns("IdCaja").Hidden = True
    
    grdCajas.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdCajas.Bands(0).Columns("Codigo").Width = 700
    
    grdCajas.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdCajas.Bands(0).Columns("Descripcion").Width = 4000
    
    grdCajas.Bands(0).Columns("ImpresoraDefault").Width = 4500
    grdCajas.Bands(0).Columns("ImpresoraDefault").Header.Caption = "Impresora Servicios"
    grdCajas.Bands(0).Columns("Impresora2").Width = 4500
    grdCajas.Bands(0).Columns("Impresora2").Header.Caption = "Impresora Farmacia"
    
    grdCajas.Bands(0).Columns("FormatoImpDefaultCinta").Hidden = True
    grdCajas.Bands(0).Columns("FormatoImp2Cinta").Hidden = True
    grdCajas.Bands(0).Columns("SerieImpresoraDefault").Hidden = True
    grdCajas.Bands(0).Columns("SerieImpresora2").Hidden = True
    
    grdCajas.Bands(0).Columns("idTipoComprobante").Hidden = True
    grdCajas.Bands(0).Columns("IdTipoComprobante2").Hidden = True
    
    'grdCajas.Bands(0).Columns("NroSerie").Header.Caption = "Nº Serie"
    'grdCajas.Bands(0).Columns("NroSerie").Width = 1000
    
    'grdCajas.Bands(0).Columns("NroComprobante").Header.Caption = "Ult.Comprobante Emitido"
    'grdCajas.Bands(0).Columns("NroComprobante").Width = 3000

End Sub
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.txtDescripcion
End Sub
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, UserControl.btnBuscar
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
     Case vbKeyF8
    End Select
       
End Sub
Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdCajas.Width = fraBusqueda.Width
   grdCajas.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub



Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdCajas, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdCajas, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub
Sub inicializar()
    SkinConfigura
End Sub



