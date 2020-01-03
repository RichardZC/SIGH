VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.UserControl ucMovimientosCaja 
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
      TabIndex        =   1
      Top             =   555
      Width           =   9930
      Begin VB.ComboBox cmbTurno 
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
         Left            =   5280
         TabIndex        =   10
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txtIdLote 
         BackColor       =   &H00FFEBD9&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   9
         Top             =   480
         Width           =   585
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
         TabIndex        =   8
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   8595
         Picture         =   "ucMovimientosCaja.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   450
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   7230
         Picture         =   "ucMovimientosCaja.ctx":2BDC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1305
      End
      Begin MSMask.MaskEdBox txtFechaLote 
         Height          =   315
         Left            =   120
         TabIndex        =   7
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
         Caption         =   "Fecha Lote           Cajero                                                       Turno           Lote"
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
         TabIndex        =   3
         Top             =   270
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
         TabIndex        =   2
         Top             =   810
         Width           =   7635
      End
   End
   Begin UltraGrid.SSUltraGrid grdComprobantes 
      Height          =   4590
      Left            =   90
      TabIndex        =   0
      Top             =   1575
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
      Caption         =   "Lista de Comprobantes"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Comprobantes de Pago"
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
      TabIndex        =   4
      Top             =   30
      Width           =   9975
   End
End
Attribute VB_Name = "ucMovimientosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MZD Ini 01/06/2005 [Todo el archivo]
'MZD02 Ini 04/07/2005

Option Explicit
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim ml_IdDepartamentoHospital As Long
Dim mo_cmbCajero As New ListaDespleglable
Dim mo_cmbTurno As New ListaDespleglable
Dim ml_IdUsuario As Long

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdComprobantes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdComprobantes.DataSource
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
Property Let IdDepartamentoHospital(lValue As Long)
    ml_IdDepartamentoHospital = lValue
End Property
Property Get IdDepartamentoHospital() As Long
    IdDepartamentoHospital = ml_IdDepartamentoHospital
End Property
Property Let IdCajero(lValue As Long)
   mo_cmbCajero.BoundText = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = Val(mo_cmbCajero.BoundText)
End Property
Property Let IdTurno(lValue As Long)
   mo_cmbTurno.BoundText = lValue
End Property
Property Get IdTurno() As Long
   IdTurno = Val(mo_cmbTurno.BoundText)
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property


Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        
    UserControl.txtIdLote.Text = ""
    Set mo_CajaLoteActual = Nothing
        
    If ((UserControl.txtFechaLote = "" Or UserControl.txtFechaLote = SIGHComun.FECHA_VACIA_DMY) Or UserControl.cmbCajero = "" Or UserControl.cmbTurno = "") Then
        MsgBox "Por favor ingrese ambos filtros (Nro Lote, Cajero, Turno)", vbInformation, "Filtro de Movimientos Caja"
        Exit Sub
    End If
    
    'Validamos la fecha
    If Not EsFecha(txtFechaLote, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, "Filtro de Movimientos Caja"
        txtFechaLote = SIGHComun.FECHA_VACIA_DMY
        Exit Sub
    End If

    Dim oCajaLote As New DOCajaLote
    oCajaLote.Fecha = UserControl.txtFechaLote.Text
    oCajaLote.IdCajero = mo_cmbCajero.BoundText
    oCajaLote.IdTurno = mo_cmbTurno.BoundText
    
    If Not mo_AdminCaja.ObtenerCajaLote(oCajaLote) Then
        MsgBox "No se ha encontrado algún lote con los datos suministrados" & vbNewLine & "Consulte con el Administrador", vbInformation, "Movimientos Caja"
        Exit Sub
    End If
    If oCajaLote.IdLote = 0 Then
        MsgBox "No se ha encontrado algún lote con los datos suministrados" & vbNewLine & "Consulte con el Administrador", vbInformation, "Movimientos Caja"
        Exit Sub
    End If
    If oCajaLote.EstadoLote = "C" Then
        MsgBox "El lote seleccionado ya está cerrado y no se pueden registrar mas comprobantes" & vbNewLine & "Consulte con el Administrador", vbInformation, "Movimientos Caja"
        Exit Sub
    End If
    
    Set mo_CajaLoteActual = oCajaLote
    txtIdLote = oCajaLote.IdLote
    
    Set grdComprobantes.DataSource = mo_AdminCaja.MovimientosCaja(oCajaLote)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox mo_AdminCaja.MensajeError, vbCritical, "Filtro órdenes de procedimientos"
    End If
    mo_Apariencia.ConfigurarFilasBiColores grdComprobantes, SIGHComun.GrillaConFilasBicolor
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtFechaLote = SIGHComun.FECHA_VACIA_DMY
        UserControl.cmbCajero = ""
End Sub

Private Sub cmbCajero_Click()
    'mo_cmbCajero.BoundColumn = "IdCajero"
    'mo_cmbCajero.ListField = "NombreCompleto"
    'Set mo_cmbCajero.RowSource = mo_AdminCaja.ServiciosSeleccionarPorTipoV2(Val(mo_cmbCajero.BoundText))
End Sub

Private Sub grdComprobantes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdComprobantes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdComprobantePago")
    

End Sub

Private Sub grdComprobantes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdComprobantes.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdComprobantePago")
    
End Sub


Private Sub grdComprobantes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdComprobantes.Bands(0).Columns("IdComprobantePago").Hidden = True
    
    grdComprobantes.Bands(0).Columns("NroSerie").Header.Caption = "Nº Serie"
    grdComprobantes.Bands(0).Columns("NroSerie").Width = 700
    
    grdComprobantes.Bands(0).Columns("NroDocumento").Header.Caption = "Nº Documento"
    grdComprobantes.Bands(0).Columns("NroDocumento").Width = 1000
    
    grdComprobantes.Bands(0).Columns("RazonSocial").Header.Caption = "Razón Social"
    grdComprobantes.Bands(0).Columns("RazonSocial").Width = 4000
    
    grdComprobantes.Bands(0).Columns("RUC").Header.Caption = "RUC"
    grdComprobantes.Bands(0).Columns("RUC").Width = 1000
    
    grdComprobantes.Bands(0).Columns("FechaCobranza").Header.Caption = "Fecha Cobranza"
    grdComprobantes.Bands(0).Columns("FechaCobranza").Width = 1200
    
    grdComprobantes.Bands(0).Columns("Observaciones").Header.Caption = "Observaciones"
    grdComprobantes.Bands(0).Columns("Observaciones").Width = 2500
    
    
    grdComprobantes.Bands(0).Columns("SubTotal").Hidden = True
    grdComprobantes.Bands(0).Columns("IGV").Hidden = True
    grdComprobantes.Bands(0).Columns("Total").Hidden = True
    grdComprobantes.Bands(0).Columns("IdLote").Hidden = True
    grdComprobantes.Bands(0).Columns("IdTipoComprobante").Hidden = True
    grdComprobantes.Bands(0).Columns("IdCuentaAtencion").Hidden = True


End Sub

Private Sub cmbCajero_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbCajero
End Sub

Private Sub cmbCajero_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtFechaLote_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaLote
End Sub

Private Sub txtFechaLote_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Public Function Inicializar()
    Set mo_cmbCajero.MiComboBox = cmbCajero
    Set mo_cmbTurno.MiComboBox = cmbTurno
    UserControl.txtFechaLote = Format(Now, SIGHComun.FormatoFechaCorta)
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
   
   grdComprobantes.Width = fraBusqueda.Width
   grdComprobantes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Sub ConfigurarCajero()
    mo_cmbCajero.ListField = "NombreCompleto"
    mo_cmbCajero.BoundColumn = "IdCajero"
    Set mo_cmbCajero.RowSource = mo_AdminCaja.CajerosSeleccionarSegunUsuario(ml_IdUsuario)
    
    mo_cmbTurno.ListField = "Descripcion"
    mo_cmbTurno.BoundColumn = "IdTurno"
    Set mo_cmbTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
    If UserControl.cmbCajero.ListCount > 0 Then
        UserControl.cmbCajero.ListIndex = 0
    End If
End Sub

Private Sub UserControl_Show()
    'ConfigurarCajero
End Sub
