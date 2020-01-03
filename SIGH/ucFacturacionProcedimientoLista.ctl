VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucFacturacionOrdenesLista 
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12690
   LockControls    =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   12690
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
      Height          =   1365
      Left            =   75
      TabIndex        =   13
      Top             =   495
      Width           =   12510
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1005
         Left            =   9120
         TabIndex        =   18
         Top             =   180
         Width           =   1875
         Begin Threed.SSOption optVentas 
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   180
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Venta Directa"
            Value           =   -1
         End
         Begin Threed.SSOption optPreventa 
            Height          =   255
            Left            =   60
            TabIndex        =   8
            Top             =   600
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "PreVenta"
         End
      End
      Begin VB.TextBox txtNcuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Width           =   1065
      End
      Begin VB.ComboBox cmbIdResponsable 
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   870
         Width           =   3915
      End
      Begin VB.TextBox txtNroOrdenPago 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         MaxLength       =   9
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbIdPtoCarga 
         Enabled         =   0   'False
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
         Left            =   6150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2925
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   11100
         Picture         =   "ucFacturacionProcedimientoLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   1275
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   11100
         Picture         =   "ucFacturacionProcedimientoLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   870
         Width           =   1275
      End
      Begin VB.TextBox txtNroOrden 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2490
         MaxLength       =   9
         TabIndex        =   12
         Top             =   480
         Width           =   885
      End
      Begin VB.TextBox txtNroHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         MaxLength       =   9
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbFarmacia 
         Height          =   405
         Left            =   11580
         TabIndex        =   17
         Top             =   60
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFinicio 
         Height          =   315
         Left            =   3390
         TabIndex        =   3
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtFfinal 
         Height          =   315
         Left            =   4770
         TabIndex        =   4
         Top             =   480
         Width           =   1350
         _ExtentX        =   2381
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
      Begin VB.Label lblNroOrden 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N°Orden                   "
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
         Left            =   2370
         TabIndex        =   20
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas                           Punto de Carga         "
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
         Left            =   4470
         TabIndex        =   19
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblServicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   930
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Cuenta    Nº Historia     "
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
         TabIndex        =   14
         Top             =   240
         Width           =   2085
      End
   End
   Begin UltraGrid.SSUltraGrid grdListaOrdenes 
      Height          =   6150
      Left            =   90
      TabIndex        =   11
      Top             =   1950
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   10848
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
      Caption         =   "Lista de ordenes"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Facturación"
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
      TabIndex        =   15
      Top             =   0
      Width           =   12585
   End
End
Attribute VB_Name = "ucFacturacionOrdenesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar Procedimientos en CONSUMO EN EL SERVICIO
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim ml_idRegistroSeleccionado As Long
Dim ml_PuntoCarga As sghTipoFiltroPacientes
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_idUsuario As Long
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim ml_IdTipoFinanciamiento As Long
Dim oRsFarmacias As New ADODB.Recordset
Dim oRsLista As New Recordset

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdListaOrdenes.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdListaOrdenes.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
    mo_cmbIdPuntoCarga.BoundText = ml_PuntoCarga
End Property
Property Get PuntoCarga() As Long
    PuntoCarga = ml_PuntoCarga
End Property
Property Let HabilitarPuntoCarga(lValue As Long)
    cmbIdPtoCarga.Enabled = lValue
End Property
Property Get HabilitarPuntoCarga() As Long
    HabilitarPuntoCarga = cmbIdPtoCarga.Enabled
End Property
Property Let idTipoFinanciamiento(lValue As Long)
    ml_IdTipoFinanciamiento = lValue
End Property
Property Get idTipoFinanciamiento() As Long
    idTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property

Private Sub btnBuscar_Click()
    If CDate(UserControl.txtFinicio.Text) > CDate(UserControl.txtFfinal.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
        Dim ldFechaIni As Date
        Dim ldFechaFin As Date
        Dim lcFiltro As String
        Dim lbSigue As Boolean
        Dim lnListIndex As Integer
        Dim rsRespuesta As New Recordset
        Dim lbServicioVacio As Boolean
        If (mo_Teclado.TextoEsSoloNumeros(UserControl.txtNcuenta.Text) = False And mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria.Text) = False And _
            UserControl.txtNroOrden = "" And mo_cmbIdPuntoCarga.BoundText = "") Then
            MsgBox "Por favor ingrese algunos de los filtros (Nro Historia, Nro cuenta, Nro Orden, Servicio)", vbInformation, "Filtro de ordenes de procedimientos"
            Exit Sub
        End If
        If mo_Teclado.TextoEsSoloNumeros(txtNroOrdenPago.Text) Then
           Set grdListaOrdenes.DataSource = mo_AdminFacturacion.FactOrdenServicioSeleccionarPorIdOrdenPago(Val(txtNroOrdenPago.Text), Val(mo_cmbIdPuntoCarga.BoundText))
         '  mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
           Exit Sub
        End If
        
        Dim oDOFactOrdenServicio As New DOFactOrdenServicio
        Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
        Dim oDOPaciente As New doPaciente
        
        Select Case ml_PuntoCarga
        Case 5  'farmacia
            
            oDOFactOrdenBienInsumo.IdOrden = Val(UserControl.txtNroOrden)
            oDOFactOrdenBienInsumo.idPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
            oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
            Set grdListaOrdenes.DataSource = mo_AdminFacturacion.AtencionOrdenesBienInsumoFiltrarDEBB(oDOFactOrdenBienInsumo, oDOPaciente, ldFechaIni, ldFechaFin, Val(cmbFarmacia.BoundText))
        
        Case Else
            lcFiltro = ""
            If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
               lcFiltro = lcFiltro & "idCuentaAtencion=" & Trim(Val(txtNcuenta.Text))
            End If
            If mo_Teclado.TextoEsSoloNumeros(txtNroOrden.Text) Then
               lcFiltro = lcFiltro & "idOrden=" & Trim(Val(txtNroOrden.Text))
            End If
            If mo_Teclado.TextoEsSoloNumeros(txtNroHistoria.Text) Then
               lcFiltro = lcFiltro & "NroHistoriaClinica=" & HCigualDNI_AgregaNUEVEaLaHistoria(txtNroHistoria.Text)
            End If
            If optVentas.Value = True Then
               If lcFiltro = "" Then
                  lcFiltro = lcFiltro & "idPaciente>0"
               Else
                  lcFiltro = lcFiltro & " and idPaciente>0"
               End If
            End If
            If optPreventa.Value = True Then
               If lcFiltro = "" Then
                  lcFiltro = lcFiltro & "idPaciente=null"
               Else
                  lcFiltro = lcFiltro & " and idPaciente=null"
               End If
            End If
            
            ldFechaIni = CDate(txtFinicio.Text & " 00:00:01")
            ldFechaFin = CDate(txtFfinal.Text & " 23:59:59")
            
            oDOFactOrdenServicio.IdOrden = Val(UserControl.txtNroOrden)
            oDOFactOrdenServicio.idPuntoCarga = Val(mo_cmbIdPuntoCarga.BoundText)
            oDOPaciente.NroHistoriaClinica = Val(UserControl.txtNroHistoria)
            If cmbIdResponsable.Visible = False Then
               lbServicioVacio = False
            Else
               lbServicioVacio = IIf(Val(mo_cmbIdResponsable.BoundText) = 0, True, False)
            End If
            lnListIndex = 0
            lbSigue = True
            Do While lbSigue = True
                If lbServicioVacio = True Then
                   cmbIdResponsable.ListIndex = lnListIndex
                   lnListIndex = lnListIndex + 1
                   If lnListIndex = cmbIdResponsable.ListCount Then
                      lbSigue = False
                   End If
                Else
                   lbSigue = False
                End If
                If cmbIdResponsable.Visible = True Then
                   Set oRsLista = mo_AdminFacturacion.FactOrdenServicioPorFechas(ldFechaIni, ldFechaFin, Val(mo_cmbIdPuntoCarga.BoundText), Val(mo_cmbIdResponsable.BoundText))
                Else
                   Set oRsLista = mo_AdminFacturacion.FactOrdenServicioPorFechas(ldFechaIni, ldFechaFin, Val(mo_cmbIdPuntoCarga.BoundText), 0)
                End If
                If lcFiltro <> "" Then
                   oRsLista.Filter = lcFiltro
                End If
                If oRsLista.RecordCount > 0 Then
                   Exit Do
                End If
            Loop
            On Error Resume Next
            Set grdListaOrdenes.DataSource = oRsLista
            If oRsLista.RecordCount = 0 Then
               MsgBox "No existe Información con esos Datos", vbInformation, "Busqueda"
               Exit Sub
            End If
            
            
        End Select
        
        If mo_AdminFacturacion.MensajeError <> "" Then
            MsgBox mo_AdminFacturacion.MensajeError, vbInformation, "Filtro órdenes de procedimientos"
        End If
        
       ' mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtNroHistoria = ""
        UserControl.txtNroOrden = ""
        UserControl.txtNroOrdenPago = ""
        UserControl.txtNcuenta.Text = ""
        mo_cmbIdResponsable.BoundText = ""
        On Error Resume Next
        UserControl.txtNcuenta.SetFocus
End Sub



Private Sub cmbIdPtoCarga_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdPtoCarga
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbIdResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsable
    AdministrarKeyPreview KeyCode

End Sub

Private Sub grdListaOrdenes_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdListaOrdenes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdOrden")
    
End Sub

Private Sub grdListaOrdenes_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdListaOrdenes.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = rsRecordset("IdOrden")
    
End Sub


Private Sub grdListaOrdenes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdListaOrdenes.Bands(0).Columns("IdPuntoCarga").Hidden = True
    grdListaOrdenes.Bands(0).Columns("IdEstadoFacturacion").Hidden = True
    
    grdListaOrdenes.Bands(0).Columns("IdOrden").Header.Caption = "Nro Orden"
    grdListaOrdenes.Bands(0).Columns("IdOrden").Width = 1200
    
    grdListaOrdenes.Bands(0).Columns("FechaDespacho").Header.Caption = "F.Despacho"
    grdListaOrdenes.Bands(0).Columns("FechaDespacho").Width = 2500
        
    grdListaOrdenes.Bands(0).Columns("idCuentaAtencion").Header.Caption = "Nro Cuenta"
    grdListaOrdenes.Bands(0).Columns("idCuentaAtencion").Width = 1200
        
    grdListaOrdenes.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "Nro Historia"
    grdListaOrdenes.Bands(0).Columns("NroHistoriaClinica").Width = 1500

    grdListaOrdenes.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Apellido Paterno"
    grdListaOrdenes.Bands(0).Columns("ApellidoPaterno").Width = 1500

    grdListaOrdenes.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Apellido Materno"
    grdListaOrdenes.Bands(0).Columns("ApellidoMaterno").Width = 1500

    grdListaOrdenes.Bands(0).Columns("PrimerNombre").Header.Caption = "Primer Nombre"
    grdListaOrdenes.Bands(0).Columns("PrimerNombre").Width = 1500

    grdListaOrdenes.Bands(0).Columns("EstadoOrden").Header.Caption = "Estado Orden"
    grdListaOrdenes.Bands(0).Columns("EstadoOrden").Width = 1500

End Sub

Private Sub grdListaOrdenes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstadoFacturacion").GetText()) = 9 Then
            Row.Appearance.ForeColor = vbRed
            'Row.Appearance.Font.Strikethrough = True
        End If

End Sub





Private Sub optPreventa_Click(Value As Integer)
    If optPreventa.Value = True Then
       lblNroOrden.Caption = "N°Ord.Pago"
       UserControl.txtNroOrden.Visible = False
       UserControl.txtNroOrdenPago.Visible = True
       btnLimpiar_Click
    End If
End Sub

Private Sub optVentas_Click(Value As Integer)
    If optVentas.Value = True Then
       lblNroOrden.Caption = "N° Orden"
       UserControl.txtNroOrden.Visible = True
       UserControl.txtNroOrdenPago.Visible = False
       btnLimpiar_Click
    End If
End Sub

Private Sub txtFfinal_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFfinal
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFfinal_LostFocus()
    If Not EsFecha(txtFfinal.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFfinal.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFinicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFinicio
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFinicio_LostFocus()
    If Not EsFecha(txtFinicio.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        txtFinicio.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtNcuenta.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub

Private Sub txtNcuenta_LostFocus()
    If txtNcuenta.Text <> "" Then
       txtNroHistoria.Text = ""
       txtNroOrdenPago.Text = ""
       txtNroOrden.Text = ""
    End If
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
   If KeyAscii = 13 And Len(txtNroHistoria.Text) > 0 Then
       btnBuscar_Click
   End If
   
End Sub

Private Sub txtNroHistoria_LostFocus()
    If txtNroHistoria.Text <> "" Then
        txtNcuenta.Text = ""
        txtNroOrdenPago.Text = ""
        txtNroOrden.Text = ""
    End If
End Sub





Private Sub txtNroOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroOrden
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroOrden_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Len(txtNroOrden.Text) > 0 Then
       btnBuscar_Click
   End If

End Sub

Private Sub txtNroOrden_LostFocus()
    If txtNroOrden.Text <> "" Then
       txtNcuenta.Text = ""
       txtNroHistoria.Text = ""
       txtNroOrdenPago.Text = ""
    End If
End Sub

Private Sub txtNroOrdenPago_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroOrdenPago
End Sub

Private Sub txtNroOrdenPago_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Len(txtNroOrdenPago.Text) > 0 Then
       btnBuscar_Click
   End If
  
End Sub

Private Sub txtNroOrdenPago_LostFocus()
    If txtNroOrdenPago.Text <> "" Then
       txtNcuenta.Text = ""
       txtNroHistoria.Text = ""
       txtNroOrden.Text = ""
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdListaOrdenes.Width = fraBusqueda.Width
   grdListaOrdenes.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighentidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdListaOrdenes, sighentidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Sub Inicializar()
    SkinConfigura
    ConfigurarPuntosDeCarga
    
    txtFinicio.Text = Date
    txtFfinal.Text = Date

    If ml_PuntoCarga = 5 Then
         cmbFarmacia.Visible = True
         lblServicio.Visible = True
         CargaFarmacias
    Else
         cmbFarmacia.Visible = False
         lblServicio.Visible = False
    End If
    '
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
    mo_cmbIdResponsable.BoundColumn = "IdServicio"
    mo_cmbIdResponsable.ListField = "DservicioHosp"
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
    Dim lcEspecialidadesDelUsuario As String
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghEspecialidadesHosp, ml_idUsuario)
    cmbIdResponsable.Visible = True: lblServicio.Visible = True
    lcEspecialidadesDelUsuario = ""
    If rsIdAlmacen.RecordCount > 0 Then
        lcEspecialidadesDelUsuario = " and ("
        rsIdAlmacen.MoveFirst
        Do While Not rsIdAlmacen.EOF
           lcEspecialidadesDelUsuario = lcEspecialidadesDelUsuario & " dbo.Servicios.idEspecialidad=" & Trim(Str(rsIdAlmacen.Fields!idLaboraSubArea)) & " or "
           rsIdAlmacen.MoveNext
        Loop
        lcEspecialidadesDelUsuario = Left(lcEspecialidadesDelUsuario, Len(lcEspecialidadesDelUsuario) - 4) & ")"
        Set mo_cmbIdResponsable.RowSource = oBuscaServicios.DevuelveServiciosDelHospital("(3)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
    Else
        Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghEspecialidadesEmergCons, ml_idUsuario)
        If rsIdAlmacen.RecordCount > 0 Then
            lcEspecialidadesDelUsuario = " and ("
            rsIdAlmacen.MoveFirst
            Do While Not rsIdAlmacen.EOF
               lcEspecialidadesDelUsuario = lcEspecialidadesDelUsuario & " dbo.Servicios.idEspecialidad=" & Trim(Str(rsIdAlmacen.Fields!idLaboraSubArea)) & " or "
               rsIdAlmacen.MoveNext
            Loop
            lcEspecialidadesDelUsuario = Left(lcEspecialidadesDelUsuario, Len(lcEspecialidadesDelUsuario) - 4) & ")"
            Set mo_cmbIdResponsable.RowSource = oBuscaServicios.DevuelveServiciosDelHospital("(2)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
        Else
           cmbIdResponsable.Visible = False: lblServicio.Visible = False
        End If
    End If
    If cmbIdResponsable.ListCount = 1 Then
       cmbIdResponsable.ListIndex = 0
    End If
    ConfiguraPermisosDelUsuario
End Sub

Sub ConfiguraPermisosDelUsuario()
    Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim oRsPermisosUsuario As New Recordset
    Set oRsPermisosUsuario = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(ml_idUsuario)
    optPreventa.Enabled = False
    optVentas.Enabled = False
    If oRsPermisosUsuario.RecordCount > 0 Then
       Do While Not oRsPermisosUsuario.EOF
          Select Case oRsPermisosUsuario.Fields!IdPermiso
          Case 116    'Facturacion - Sólo realiza PreVenta de Servicios
               optPreventa.Enabled = True
               optPreventa.Value = True
          Case 117    'Facturacion - Sólo realiza VentaDirecta de Servicios
               optVentas.Enabled = True
               optVentas.Value = True
          End Select
          oRsPermisosUsuario.MoveNext
       Loop
    End If
    Set oRsPermisosUsuario = Nothing
End Sub


Sub CargaFarmacias()
        On Error GoTo ErrFarm
        Dim oConexion As New ADODB.Connection
        Dim lnCodigoFarmacia  As Long
        Dim lcSql As String
        
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        Set oRsFarmacias = mo_ReglasDeSeguridad.UsuariosRolesXidEmpleadoEsDeFarmacia(ml_idUsuario, oConexion)
        If oRsFarmacias.RecordCount > 0 Then
            lnCodigoFarmacia = oRsFarmacias.Fields!IdPermiso
        End If
        oRsFarmacias.Close
        Set oRsFarmacias = mo_ReglasDeSeguridad.PermisosSoloFarmacia(oConexion)
        Set cmbFarmacia.RowSource = oRsFarmacias
        cmbFarmacia.ListField = "descripcion"
        cmbFarmacia.BoundColumn = "idPermiso"
        If lnCodigoFarmacia > 0 Then
           cmbFarmacia.BoundText = lnCodigoFarmacia
        End If
ErrFarm:
'         cmbFarmacia.BoundText = ""
'         cmbFarmacia.Text = ""
'        oRsFarmacias.Close
'        Resume
End Sub

Sub ConfigurarPuntosDeCarga()
    
    Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPtoCarga
    mo_cmbIdPuntoCarga.ListField = "Descripcion"
    mo_cmbIdPuntoCarga.BoundColumn = "IdPuntoCarga"
    
    Set mo_cmbIdPuntoCarga.RowSource = mo_reglasComunes.SeleccionarPuntosDeCarga()

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
