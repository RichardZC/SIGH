VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form OrdenesPendientesPagoBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12615
   Icon            =   "OrdenesPendientesPagoBusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabFarmServ 
      Height          =   4515
      Left            =   30
      TabIndex        =   12
      Top             =   960
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   7964
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Farmacia"
      TabPicture(0)   =   "OrdenesPendientesPagoBusqueda.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdAdmision"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Servicio"
      TabPicture(1)   =   "OrdenesPendientesPagoBusqueda.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdTab1"
      Tab(1).ControlCount=   1
      Begin UltraGrid.SSUltraGrid grdAdmision 
         Height          =   3930
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   6932
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdAdmision"
      End
      Begin UltraGrid.SSUltraGrid grdTab1 
         Height          =   3930
         Left            =   -74880
         TabIndex        =   14
         Top             =   420
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   6932
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdTab1"
      End
   End
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
      Height          =   885
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   12525
      Begin VB.CommandButton cmdSinApellidoMaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   6000
         TabIndex        =   17
         ToolTipText     =   "Sin apellido MATERNO"
         Top             =   450
         Width           =   255
      End
      Begin VB.CommandButton cmdSinApellidoPaterno 
         Caption         =   "..."
         Height          =   315
         Left            =   4380
         TabIndex        =   16
         ToolTipText     =   "Sin apellido PATERNO"
         Top             =   450
         Width           =   255
      End
      Begin VB.TextBox txtDni 
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
         Left            =   90
         MaxLength       =   8
         TabIndex        =   15
         Top             =   450
         Width           =   1395
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
         Left            =   4800
         MaxLength       =   40
         TabIndex        =   2
         Top             =   450
         Width           =   1185
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
         Left            =   1590
         MaxLength       =   9
         TabIndex        =   0
         Top             =   450
         Width           =   1395
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   11100
         Picture         =   "OrdenesPendientesPagoBusqueda.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   9720
         Picture         =   "OrdenesPendientesPagoBusqueda.frx":38DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   420
         Width           =   1305
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
         Left            =   3090
         MaxLength       =   40
         TabIndex        =   1
         Top             =   450
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtFecha1 
         Height          =   315
         Left            =   6360
         TabIndex        =   3
         Top             =   450
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
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   7800
         TabIndex        =   4
         Top             =   450
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
         Caption         =   "DNI                    Historia clínica        Apellido paterno    Apellido materno            Fechas Orden Pago"
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
         TabIndex        =   11
         Top             =   225
         Width           =   8715
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   45
      TabIndex        =   9
      Top             =   5505
      Width           =   12510
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "OrdenesPendientesPagoBusqueda.frx":6527
         DownPicture     =   "OrdenesPendientesPagoBusqueda.frx":69EB
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
         Left            =   6397
         Picture         =   "OrdenesPendientesPagoBusqueda.frx":6ED7
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "OrdenesPendientesPagoBusqueda.frx":73C3
         DownPicture     =   "OrdenesPendientesPagoBusqueda.frx":7823
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
         Left            =   4852
         Picture         =   "OrdenesPendientesPagoBusqueda.frx":7C98
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "OrdenesPendientesPagoBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Ordenes de Farmacia y Servicio sin pago en CAJA
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_TipoProducto As Integer
Dim ml_idOrdenSeleccionado As Long
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mb_lbEstoyEnTabServicio As Boolean

Property Get lbEstoyEnTabServicio() As Boolean
   lbEstoyEnTabServicio = mb_lbEstoyEnTabServicio
End Property


Property Let idOrdenSeleccionado(lValue As Long)
   ml_idOrdenSeleccionado = lValue
End Property
Property Get idOrdenSeleccionado() As Long
   idOrdenSeleccionado = ml_idOrdenSeleccionado
End Property
Property Let TipoProducto(lValue As Long)
   ml_TipoProducto = lValue
End Property
Property Get TipoProducto() As Long
   TipoProducto = ml_TipoProducto
End Property


Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
   grdAdmision_DblClick
End Sub

Private Sub btnBuscar_Click()
    BusquedaDatos
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub


Private Sub btnLimpiar_Click()
    txtFecha1.Text = Date
    txtFecha2.Text = Date
    txtApellidoPaterno.Text = ""
    txtApellidoMaterno.Text = ""
    txtNroHistoria.Text = ""
End Sub

Private Sub cmdSinApellidoMaterno_Click()
    txtApellidoMaterno.Text = wxSinApellido
End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    mb_lbEstoyEnTabServicio = False
    txtFecha1.Text = Date
    txtFecha2.Text = Date
    If ml_TipoProducto <> sghAmbos Then
       Me.tabFarmServ.TabVisible(1) = False
       Me.tabFarmServ.TabCaption(0) = ""
    End If
    BusquedaDatos
    mo_Apariencia.ConfigurarFilasBiColores grdAdmision, sighentidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores Me.grdTab1, sighentidades.GrillaConFilasBicolor
End Sub

Sub BusquedaDatos()
    Dim lcSql  As String
    Dim oRsTmp As New ADODB.Recordset
    Dim oRsTmpTab1 As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim oBuscaPendientes As New Recordset
    Dim ldFecha1 As Date, ldFecha2 As Date
    oConexion.CommandTimeout = 150
    oConexion.Open sighentidades.CadenaConexion
    If ml_TipoProducto = sghbien Then
      Set oRsTmp = mo_AdminCaja.SeleccionarOrdenesPendientesDePago(CDate(txtFecha1.Text), CDate(txtFecha2.Text))
    ElseIf ml_TipoProducto = sghServicio Then
      Set oRsTmp = mo_AdminCaja.FactOrdenServicioPagosPendientesDePagoPorFechas(CDate(txtFecha1.Text), CDate(txtFecha2.Text) + 1)
    Else
      Set oRsTmp = mo_AdminCaja.SeleccionarOrdenesPendientesDePago(CDate(txtFecha1.Text), CDate(txtFecha2.Text))
      Set oRsTmpTab1 = mo_AdminCaja.FactOrdenServicioPagosPendientesDePagoPorFechas(CDate(txtFecha1.Text), CDate(txtFecha2.Text) + 1)
    End If
    lcSql = ""
    If mo_Teclado.TextoEsSoloNumeros(txtNroHistoria.Text) Then
      lcSql = lcSql & "NroHistoriaClinica=" & HCigualDNI_AgregaNUEVEaLaHistoria(txtNroHistoria.Text) & " and "
    ElseIf txtDni.Text <> "" Then
      lcSql = lcSql & "NroDocumento='" & txtDni.Text & "' and idDocIdentidad=1 and "
    Else
      If txtApellidoPaterno.Text <> "" Then
         lcSql = lcSql & "ApellidoPaterno like '" & Trim(txtApellidoPaterno.Text) & "%' and "
      End If
      If txtApellidoMaterno.Text <> "" Then
         lcSql = lcSql & "ApellidoMaterno like '" & Trim(txtApellidoMaterno.Text) & "%' and "
      End If
    End If
    If lcSql <> "" Then
       lcSql = Left(lcSql, Len(lcSql) - 5)
       oRsTmp.Filter = lcSql
    End If
    If ml_TipoProducto <> sghAmbos Then
       Set grdAdmision.DataSource = oRsTmp
    Else
       If lcSql <> "" Then
          oRsTmpTab1.Filter = lcSql
       End If
       Set grdAdmision.DataSource = oRsTmp
       Set grdTab1.DataSource = oRsTmpTab1
    End If
    
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



Private Sub grdAdmision_DblClick()
    Dim oRsTmp As New ADODB.Recordset
    Set oRsTmp = grdAdmision.DataSource
    If oRsTmp.RecordCount > 0 Then
        If ml_TipoProducto = sghbien Then
            ml_idOrdenSeleccionado = IIf(IsNull(oRsTmp.Fields!Nro_Preventa), 0, oRsTmp.Fields!Nro_Preventa)
        ElseIf ml_TipoProducto = sghServicio Then
            ml_idOrdenSeleccionado = oRsTmp.Fields!IdOrdenPago
        Else
            ml_idOrdenSeleccionado = IIf(IsNull(oRsTmp.Fields!Nro_Preventa), 0, oRsTmp.Fields!Nro_Preventa)
        End If
        mi_BotonPresionado = sghAceptar
        Me.Visible = False
    End If
End Sub

Private Sub grdAdmision_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
End Sub

Private Sub grdAdmision_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
        grdAdmision_DblClick
    End If
End Sub


Private Sub grdTab1_DblClick()
    Dim oRsTmp As New ADODB.Recordset
    Set oRsTmp = Me.grdTab1.DataSource
    If oRsTmp.RecordCount > 0 Then
        ml_idOrdenSeleccionado = oRsTmp.Fields!IdOrdenPago
        mi_BotonPresionado = sghAceptar
        Me.Visible = False
    End If
End Sub



Private Sub grdTab1_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
End Sub

Private Sub grdTab1_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
        grdTab1_DblClick
    End If

End Sub

Private Sub tabFarmServ_Click(PreviousTab As Integer)
        If tabFarmServ.Tab = 1 Then
           mb_lbEstoyEnTabServicio = True
        Else
           mb_lbEstoyEnTabServicio = False
        End If
End Sub


Private Sub txtDni_LostFocus()
    If Len(txtDni.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub



Private Sub txtFecha1_LostFocus()
If Not EsFecha(txtFecha1.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFecha1.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFecha2_LostFocus()
If Not EsFecha(txtFecha2.Text, "DD/MM/AAAA") Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFecha2.Text = sighentidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtNroHistoria_LostFocus()
    If Len(txtNroHistoria.Text) > 0 Then
       btnBuscar_Click
    End If
End Sub
