VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucMovimientoFormatoHcLista 
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   LockControls    =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   10200
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      Height          =   1515
      Left            =   75
      TabIndex        =   8
      Top             =   555
      Width           =   10035
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9225
         Picture         =   "ucMovimientoFormatoHcListactl.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7860
         Picture         =   "ucMovimientoFormatoHcListactl.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtNroHistoria 
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
         Left            =   150
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1725
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   2100
         MaxLength       =   40
         TabIndex        =   1
         Top             =   480
         Width           =   1845
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   4035
         MaxLength       =   40
         TabIndex        =   2
         Top             =   480
         Width           =   1845
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   5970
         MaxLength       =   40
         TabIndex        =   3
         Top             =   480
         Width           =   1845
      End
      Begin MSMask.MaskEdBox txtFechaMovimiento 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   1035
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "  Nº historia clínica           Apellido paterno           Apellido materno        Primer nombre                   "
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
         TabIndex        =   10
         Top             =   240
         Width           =   7815
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha movimiento"
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
         TabIndex        =   9
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdMovimientos 
      Height          =   4230
      Left            =   75
      TabIndex        =   7
      Top             =   2160
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7461
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de movimientos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Movimiento de Formatos de Historia Clínica"
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
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   10200
   End
End
Attribute VB_Name = "ucMovimientoFormatoHcLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista Movimientos de Formatos de Historia
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria
Dim mo_Teclado As New sighEntidades.Teclado

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdMovimientos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdMovimientos.DataSource
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
Property Let TipoBusqueda(lValue As sghTipoBusquedaPrestamoHistoria)
    ml_TipoBusqueda = lValue
End Property
Property Get TipoBusqueda() As sghTipoBusquedaPrestamoHistoria
    TipoBusqueda = ml_TipoBusqueda
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub
Public Sub RealizarBusqueda()
Dim oPaciente As New doPaciente
Dim oMovimiento As New DOMovimientoHistoriaClinica
        
        If (UserControl.txtApellidoPaterno = "" And UserControl.txtApellidoMaterno = "" And _
            UserControl.txtPrimerNombre = "" And UserControl.txtNroHistoria = "" And UserControl.txtFechaMovimiento = sighEntidades.FECHA_VACIA_DMY) Then
            MsgBox "Ingrese algunos de los filtro para realizar la búsqueda", vbInformation, "Movimiento de Historias"
            Exit Sub
        End If
            
        
        oPaciente.ApellidoMaterno = UserControl.txtApellidoMaterno
        oPaciente.ApellidoPaterno = UserControl.txtApellidoPaterno
        oPaciente.PrimerNombre = UserControl.txtPrimerNombre
        oPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
        oMovimiento.FechaMovimiento = IIf(UserControl.txtFechaMovimiento = sighEntidades.FECHA_VACIA_DMY, 0, UserControl.txtFechaMovimiento)
        
        Set grdMovimientos.DataSource = mo_AdminArchivoClinico.MovimientosFormatosHCFiltrar(oPaciente, oMovimiento)
        
        If mo_AdminArchivoClinico.MensajeError <> "" Then
            MsgBox mo_AdminArchivoClinico.MensajeError, vbInformation, "Filtro PrestamosHC"
        End If
        
        mo_Apariencia.ConfigurarFilasBiColores grdMovimientos, sighEntidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtApellidoMaterno = ""
        UserControl.txtApellidoPaterno = ""
        UserControl.txtPrimerNombre = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtFechaMovimiento = sighEntidades.FECHA_VACIA_DMY
End Sub
Private Sub grdMovimientos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset
    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdMovimientos.DataSource
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
End Sub

Private Sub grdMovimientos_Click()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdMovimientos.DataSource
    ml_idRegistroSeleccionado = rsRecordset("IdMovimiento")
        
End Sub

Private Sub grdMovimientos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    
    grdMovimientos.Bands(0).Columns("IdMovimiento").Header.Caption = "Id"
    grdMovimientos.Bands(0).Columns("IdMovimiento").Width = 1000
    
    grdMovimientos.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nro Historia"
    grdMovimientos.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    grdMovimientos.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdMovimientos.Bands(0).Columns("ApellidoPaterno").Width = 1200
    
    grdMovimientos.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdMovimientos.Bands(0).Columns("ApellidoMaterno").Width = 1200
    
    grdMovimientos.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdMovimientos.Bands(0).Columns("PrimerNombre").Width = 1200
    
    grdMovimientos.Bands(0).Columns("FechaMovimiento").Header.Caption = "F.Movimiento"
    grdMovimientos.Bands(0).Columns("FechaMovimiento").Width = 1400
    grdMovimientos.Bands(0).Columns("FechaMovimiento").Format = sighEntidades.DevuelveFechaSoloFormato_DMY_HM
    
    grdMovimientos.Bands(0).Columns("Origen").Header.Caption = "Origen."
    grdMovimientos.Bands(0).Columns("Origen").Width = 2000
    
    grdMovimientos.Bands(0).Columns("Destino").Header.Caption = "Destino."
    grdMovimientos.Bands(0).Columns("Destino").Width = 2000

    grdMovimientos.Bands(0).Columns("Observacion").Header.Caption = "Observación"
    grdMovimientos.Bands(0).Columns("Observacion").Width = 2500
    
    grdMovimientos.Bands(0).Columns("NroFolios").Header.Caption = "Nº Folio"
    grdMovimientos.Bands(0).Columns("NroFolios").Width = 1000
    
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdMovimientos.Width = fraBusqueda.Width
   grdMovimientos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Private Sub txtFechaMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaMovimiento
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaMovimiento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
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
         btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub

Public Sub inicializar()
    txtFechaMovimiento.Text = Date
End Sub
