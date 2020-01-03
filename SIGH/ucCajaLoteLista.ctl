VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucCajaLoteLista 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   ScaleHeight     =   6210
   ScaleWidth      =   10065
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
      Left            =   45
      TabIndex        =   0
      Top             =   525
      Width           =   9930
      Begin VB.ComboBox cmbCaja 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7170
         Picture         =   "ucCajaLoteLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8535
         Picture         =   "ucCajaLoteLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1215
      End
      Begin VB.ComboBox cmbCajero 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   3315
      End
      Begin MSMask.MaskEdBox txtFechaLote 
         Height          =   315
         Left            =   120
         TabIndex        =   4
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha               Cajero                                                Caja"
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
         TabIndex        =   5
         Top             =   270
         Width           =   6975
      End
   End
   Begin UltraGrid.SSUltraGrid grdLotes 
      Height          =   4590
      Left            =   60
      TabIndex        =   7
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
      Caption         =   "Lista de Lotes"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Asignaciones de Caja"
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
      TabIndex        =   8
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "ucCajaLoteLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MZD Ini 19/06/2005 [Todo el archivo]

Option Explicit
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_cmbCajero As New ListaDespleglable
Dim mo_cmbCaja As New ListaDespleglable

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdLotes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdLotes.DataSource
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
    Dim oLote As New DOCajaLote
    
    'Validamos la fecha
    If txtFechaLote.Text <> "" And txtFechaLote.Text <> SIGHComun.FECHA_VACIA_DMY Then
        If Not EsFecha(txtFechaLote, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, "Filtro de Lotes"
            txtFechaLote = SIGHComun.FECHA_VACIA_DMY
            Exit Sub
        End If
        oLote.Fecha = txtFechaLote.Text
    Else
        oLote.Fecha = 0
    End If
    oLote.IdCajero = Val(mo_cmbCajero.BoundText)
    oLote.IdCaja = Val(mo_cmbCaja.BoundText)
        
    Set grdLotes.DataSource = mo_AdminCaja.RealizarFiltroLotes(oLote)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox mo_AdminCaja.MensajeError, vbCritical, "Filtro de supervisores"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdLotes, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
    UserControl.txtFechaLote = SIGHComun.FECHA_VACIA_DMY
    mo_cmbCajero.BoundText = ""
    mo_cmbCaja.BoundText = ""
End Sub

Private Sub grdLotes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdLotes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdLote")

End Sub

Private Sub grdLotes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdLotes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdLote")
    
End Sub


Private Sub grdLotes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdLotes.Bands(0).Columns("IdLote").Hidden = True
    
    grdLotes.Bands(0).Columns("NombreCaja").Header.Caption = "Caja"
    grdLotes.Bands(0).Columns("NombreCaja").Width = 2000
    
    'MZD02 Ini 04/07/2005
    grdLotes.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdLotes.Bands(0).Columns("Turno").Width = 2000
    'MZD02 Fin 04/07/2005
    
    grdLotes.Bands(0).Columns("NombreCajero").Header.Caption = "Cajero"
    grdLotes.Bands(0).Columns("NombreCajero").Width = 3000
    
    grdLotes.Bands(0).Columns("Fecha").Header.Caption = "Fecha"
    grdLotes.Bands(0).Columns("Fecha").Width = 1200
    
    grdLotes.Bands(0).Columns("EstadoLote").Header.Caption = "Estado Lote"
    grdLotes.Bands(0).Columns("EstadoLote").Width = 1500

    grdLotes.Bands(0).Columns("SaldoInicialSoles").Header.Caption = "Saldo Ini.(S/.)"
    grdLotes.Bands(0).Columns("SaldoInicialSoles").Width = 1500

    grdLotes.Bands(0).Columns("SaldoInicialDolares").Header.Caption = "Saldo Ini.($)"
    grdLotes.Bands(0).Columns("SaldoInicialDolares").Width = 1500
End Sub

Public Function Inicializar()
    Set mo_cmbCaja.MiComboBox = UserControl.cmbCaja
    Set mo_cmbCajero.MiComboBox = UserControl.cmbCajero
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
   
   grdLotes.Width = fraBusqueda.Width
   grdLotes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub
Private Sub CargarComboBoxes()
    mo_cmbCaja.BoundColumn = "IdCaja"
    mo_cmbCaja.ListField = "Descripcion"
    Set mo_cmbCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()

    mo_cmbCajero.BoundColumn = "IdCajero"
    mo_cmbCajero.ListField = "NombreCompleto"
    Set mo_cmbCajero.RowSource = mo_AdminCaja.CajerosSeleccionarTodosParaLista()
End Sub


