VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucCartaGarantiaLista 
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10050
   ScaleHeight     =   6225
   ScaleWidth      =   10050
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
      Left            =   60
      TabIndex        =   3
      Top             =   525
      Width           =   9930
      Begin VB.TextBox txtNroCuenta 
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
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   1
         Top             =   480
         Width           =   1845
      End
      Begin VB.TextBox txtNroCarta 
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
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   1845
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7200
         Picture         =   "ucCartaGarantiaLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5820
         Picture         =   "ucCartaGarantiaLista.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFechaVigencia 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Vigencia     Cuenta                       Nro Carta"
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
         Top             =   270
         Width           =   5415
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
   Begin UltraGrid.SSUltraGrid grdCartaGarantia 
      Height          =   4590
      Left            =   60
      TabIndex        =   8
      Top             =   1545
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   8096
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
      Caption         =   "Lista de Cartas de Garantía"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Cartas de Garantía"
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
      TabIndex        =   9
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucCartaGarantiaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Apariencia As New SIGHComun.GridInfragistic

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdCartaGarantia.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdCartaGarantia.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
    Dim oDOCartaGarantia As New DOCartaGarantia
    
    'Validamos la fecha
    If txtFechaVigencia.Text <> "" And txtFechaVigencia.Text <> SIGHComun.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaVigencia, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, "Filtro de Cartas de Garantía"
            txtFechaVigencia = SIGHComun.FECHA_VACIA_DMY
            Exit Sub
        End If
        oDOCartaGarantia.FechaVigencia = txtFechaVigencia.Text
    Else
        oDOCartaGarantia.FechaVigencia = 0
    End If
    oDOCartaGarantia.IdCuentaAtencion = Val(UserControl.txtNroCuenta.Text)
    oDOCartaGarantia.NroCarta = UserControl.txtNroCarta.Text
        
    Set grdCartaGarantia.DataSource = mo_AdminComun.CartaGarantiaFiltrar(oDOCartaGarantia)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox mo_AdminComun.MensajeError, vbCritical, "Filtro de Cartas de Garantía"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdCartaGarantia, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtFechaVigencia = SIGHComun.FECHA_VACIA_DMY
    UserControl.txtNroCarta = ""
    UserControl.txtNroCuenta = ""
End Sub

Private Sub grdCartaGarantia_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCartaGarantia.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdCartaGarantia")

End Sub

Private Sub grdCartaGarantia_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCartaGarantia.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdCartaGarantia")
    
End Sub


Private Sub grdCartaGarantia_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdCartaGarantia.Bands(0).Columns("IdCartaGarantia").Hidden = True
    
    grdCartaGarantia.Bands(0).Columns("ValorCobertura").Header.Caption = "Cobertura"
    grdCartaGarantia.Bands(0).Columns("ValorCobertura").Width = 2000
    
    grdCartaGarantia.Bands(0).Columns("Observacion").Header.Caption = "Observación"
    grdCartaGarantia.Bands(0).Columns("Observacion").Width = 2000
    
    grdCartaGarantia.Bands(0).Columns("NroCarta").Header.Caption = "Nº Carta"
    grdCartaGarantia.Bands(0).Columns("NroCarta").Width = 2000
    
    grdCartaGarantia.Bands(0).Columns("FechaVigencia").Header.Caption = "Fecha de Vigencia"
    grdCartaGarantia.Bands(0).Columns("FechaVigencia").Width = 2000
    
    grdCartaGarantia.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "Nº Cuenta Atencion"
    grdCartaGarantia.Bands(0).Columns("IdCuentaAtencion").Width = 2000
    
End Sub

Private Sub UserControl_Initialize()
    'CargarComboBoxes
End Sub

Public Function Inicializar()
    CargarComboBoxes
End Function

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
   
   grdCartaGarantia.Width = fraBusqueda.Width
   grdCartaGarantia.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub
Private Sub CargarComboBoxes()

End Sub



