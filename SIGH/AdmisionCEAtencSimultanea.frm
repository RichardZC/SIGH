VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form AdmisionCEatencSimultanea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultorios que registran información al mismo tiempo"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   15480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraAgregar 
      Caption         =   "Agregar"
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
      Height          =   2820
      Index           =   0
      Left            =   11865
      TabIndex        =   8
      Top             =   3675
      Visible         =   0   'False
      Width           =   3555
      Begin VB.TextBox txtEdadFin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   18
         Top             =   945
         Width           =   465
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   105
         TabIndex        =   28
         Top             =   2460
         Width           =   840
      End
      Begin VB.Frame Frame 
         Height          =   825
         Index           =   1
         Left            =   105
         TabIndex        =   24
         Top             =   1590
         Width           =   3360
         Begin VB.CommandButton cmdBuscaCptDx 
            Caption         =   "..."
            Height          =   285
            Left            =   195
            TabIndex        =   22
            Top             =   465
            Width           =   390
         End
         Begin Threed.SSOption optCpt 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   135
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Caption         =   "Cpt"
            Value           =   -1
         End
         Begin Threed.SSOption optDx 
            Height          =   315
            Left            =   1815
            TabIndex        =   26
            Top             =   135
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
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
            Caption         =   "Dx"
         End
         Begin VB.Label lblDxCpt 
            AutoSize        =   -1  'True
            Caption         =   "..."
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
            Left            =   630
            TabIndex        =   27
            Top             =   495
            Width           =   180
         End
      End
      Begin VB.ComboBox cmdTipoEdad 
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
         ItemData        =   "AdmisionCEAtencSimultanea.frx":0000
         Left            =   2235
         List            =   "AdmisionCEAtencSimultanea.frx":000D
         TabIndex        =   19
         Top             =   930
         Width           =   1230
      End
      Begin VB.TextBox txtPesoFinal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2940
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1260
         Width           =   495
      End
      Begin VB.TextBox txtPesoInicial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   930
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1275
         Width           =   465
      End
      Begin VB.TextBox txtEdad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   915
         MaxLength       =   4
         TabIndex        =   17
         Top             =   945
         Width           =   465
      End
      Begin VB.TextBox txtLab 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   915
         MaxLength       =   3
         TabIndex        =   16
         Top             =   615
         Width           =   465
      End
      Begin VB.TextBox txtSubGrupo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2955
         MaxLength       =   2
         TabIndex        =   15
         Top             =   255
         Width           =   465
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
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
         Left            =   2580
         TabIndex        =   23
         Top             =   2475
         Width           =   840
      End
      Begin VB.TextBox txtGrupo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   915
         MaxLength       =   2
         TabIndex        =   14
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   1560
         TabIndex        =   30
         Top             =   990
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   2460
         TabIndex        =   29
         Top             =   1305
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Peso (Kg)"
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
         Left            =   90
         TabIndex        =   13
         Top             =   1335
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Edad"
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
         Left            =   90
         TabIndex        =   12
         Top             =   985
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Lab"
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
         Left            =   90
         TabIndex        =   11
         Top             =   635
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SubGrupo"
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
         Left            =   2100
         TabIndex        =   10
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   90
         TabIndex        =   9
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame FraConsideraciones 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2385
      Left            =   13290
      TabIndex        =   7
      Top             =   75
      Visible         =   0   'False
      Width           =   2160
      Begin VB.CommandButton cmdFiltraActividad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   180
         Picture         =   "AdmisionCEAtencSimultanea.frx":0024
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Filtra actividades"
         Top             =   975
         Width           =   1845
      End
      Begin VB.TextBox txtEdad11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   645
         MaxLength       =   4
         TabIndex        =   33
         Text            =   "1"
         Top             =   240
         Width           =   465
      End
      Begin VB.ComboBox cmbTipoEd11 
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
         ItemData        =   "AdmisionCEAtencSimultanea.frx":05AE
         Left            =   180
         List            =   "AdmisionCEAtencSimultanea.frx":05BB
         TabIndex        =   32
         Top             =   570
         Width           =   1875
      End
      Begin VB.TextBox txtEdad12 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "120"
         Top             =   225
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Edad"
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
         Left            =   210
         TabIndex        =   35
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   1380
         TabIndex        =   34
         Top             =   270
         Width           =   120
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "AdmisionCEAtencSimultanea.frx":05D2
      DownPicture     =   "AdmisionCEAtencSimultanea.frx":0A96
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
      Left            =   11880
      Picture         =   "AdmisionCEAtencSimultanea.frx":0F82
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1725
      Width           =   1365
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
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
      Left            =   11880
      TabIndex        =   5
      Top             =   960
      Width           =   1365
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Agregar"
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
      Left            =   11880
      TabIndex        =   4
      ToolTipText     =   "New"
      Top             =   180
      Width           =   1365
   End
   Begin VB.Frame fraAceptar 
      Height          =   1035
      Left            =   45
      TabIndex        =   3
      Top             =   5415
      Width           =   11775
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AdmisionCEAtencSimultanea.frx":146E
         DownPicture     =   "AdmisionCEAtencSimultanea.frx":18CE
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
         Left            =   4470
         Picture         =   "AdmisionCEAtencSimultanea.frx":1D43
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AdmisionCEAtencSimultanea.frx":21B8
         DownPicture     =   "AdmisionCEAtencSimultanea.frx":267C
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
         Left            =   6037
         Picture         =   "AdmisionCEAtencSimultanea.frx":2B68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdBusqueda 
      Height          =   5400
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   9525
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
      Caption         =   "Lista de UPS"
   End
End
Attribute VB_Name = "AdmisionCEatencSimultanea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Lista atenciones de un Paciente
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idCuentaAtencion  As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_ReglasAdmision  As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_Facturacion As New SIGHNegocios.ReglasFacturacion
'Dim oRsFuas As New Recordset
Dim ml_EstadoCuenta As String
Dim ml_idCorrelativo As Long
Dim ml_formLlamante As String       'boton agregar cpt desde atencion Medico, boton FUA desde atencion del MEdico
Dim ml_NroFua As Long
Dim ml_FuaIdCuentaAtencion As Long
Dim mrs_oRsFua As Recordset
Dim oRsItemsMasivosElegidos As New Recordset
Dim oRsTipoDx As New Recordset
Dim oRsActividades As New Recordset
Dim lbPulsoBotonAceptar As Boolean
Dim ml_ups As String
Dim lcSql As String

Property Let UPS(lValue As String)
    ml_ups = lValue
End Property

Property Set oRsItemsElegidos(oValue As Recordset)
    Set oRsTipoDx = oValue
End Property

Property Set oRsFua(oValue As Recordset)
    Set mrs_oRsFua = oValue
End Property
Property Get NroFua() As Long
   NroFua = ml_NroFua
End Property
Property Get FuaIdCuentaAtencion() As Long
   FuaIdCuentaAtencion = ml_FuaIdCuentaAtencion
End Property
Property Let FormLlamante(lValue As String)
    ml_formLlamante = lValue
End Property
Property Let Correlativo(lValue As Long)
    ml_idCorrelativo = lValue
End Property
Property Let EstadoCuenta(lValue As String)
    ml_EstadoCuenta = lValue
End Property
Property Get EstadoCuenta() As String
    EstadoCuenta = ml_EstadoCuenta
End Property

Property Set Atenciones(oValue As Recordset)
    Set grdBusqueda.DataSource = oValue
End Property

Property Get Atenciones() As Recordset
End Property

Property Let idCuentaAtencion(lValue As Long)
    ml_idCuentaAtencion = lValue
End Property

Property Get idCuentaAtencion() As Long
    idCuentaAtencion = ml_idCuentaAtencion
End Property

Private Sub btnAceptar_Click()
    grdBusqueda_DblClick
    lbPulsoBotonAceptar = True
End Sub

Private Sub btnCancelar_Click()
    ml_idCuentaAtencion = 0
    Select Case ml_formLlamante
    Case "ACTIVIDADES"
         lbPulsoBotonAceptar = False
    End Select
    Me.Visible = False
End Sub



Private Sub cmdBuscaCptDx_Click()
    Me.lblDxCpt.Caption = ""
    Me.lblDxCpt.Tag = ""
    If Me.optDx.Value = True Then
        Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
        Dim oDODiagnostico As DODiagnostico
        oBusqueda.SoloMuestraDxGalenHos = False
        oBusqueda.MostrarFormulario
        If oBusqueda.BotonPresionado = sghAceptar Then
            Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
            If Not oDODiagnostico Is Nothing Then
                Me.lblDxCpt.Tag = oDODiagnostico.CodigoCie2004
                Me.lblDxCpt.Caption = oDODiagnostico.Descripcion
            End If
        End If
        Set oBusqueda = Nothing
    Else
        Dim oFrm As New SIGHNegocios.BuscaServicio
        Dim dOServ As New DOCatalogoServicio
        oFrm.MostrarFormulario
        If oFrm.IdRegistroSeleccionado <> 0 Then
            Set dOServ = mo_Facturacion.CatalogoServiciosSeleccionarPorId(oFrm.IdRegistroSeleccionado)
            If Not dOServ Is Nothing Then
                Me.lblDxCpt.Tag = dOServ.codigo
                Me.lblDxCpt.Caption = dOServ.nombre
            End If
        End If
        Set oFrm = Nothing
        Set dOServ = Nothing
    End If
End Sub

Private Sub cmdCancelar_Click()
mo_Formulario.HabilitarDeshabilitar FraAgregar(0), False
LimpiaDatos
End Sub

Private Sub cmdDel_Click()
   If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        oRsActividades.Delete
        oRsActividades.Update
        oRsActividades.Requery
    End If
End Sub



Private Sub cmdFiltraActividad_Click()
         lcSql = "select * from ServiciosAtenSimultaneaImpHIS where ups='" & ml_ups & _
         "' and idTipoEdad=" & IIf(Me.cmbTipoEd11.ListIndex = 0, "1", IIf(Me.cmbTipoEd11.ListIndex = 1, "2", "3")) & _
         "  and edadInicio>=" & Me.txtEdad11.Text & " and edadFinal<=" & Me.txtEdad12.Text & _
         " order by grupo, subgrupoorden"
         If oRsActividades.State = 1 Then oRsActividades.Close
         oRsActividades.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
         Set grdBusqueda.DataSource = oRsActividades

End Sub

Private Sub cmdGrabar_Click()
    Dim lbNuevo As Boolean
    If Val(Me.txtGrupo.Text) <= 0 Then
       MsgBox "Grupo debe ser mayor a CERO", vbInformation, ""
       Exit Sub
    End If
    If Val(Me.txtSubGrupo.Text) <= 0 Then
       MsgBox "SubGrupo debe ser mayor a CERO", vbInformation, ""
       Exit Sub
    End If
    If Me.lblDxCpt.Caption = "" Then
       MsgBox "Debe elegir Dx/Cpt", vbInformation, ""
       Exit Sub
    End If
    If Me.cmdTipoEdad.Text = "" Then
       MsgBox "Elija TIPO EDAD", vbInformation, ""
       Exit Sub
    End If
    If Val(Me.txtEdad.Text) > Val(Me.txtEdadFin.Text) Then
       MsgBox "La edad FINAL debe ser mayor a la edad INICIAL", vbInformation, ""
       Exit Sub
    End If
    If Val(Me.txtPesoInicial.Text) > Val(Me.txtPesoFinal.Text) Then
       MsgBox "El peso FINAL no puede ser menor al peso INICIAL", vbInformation, ""
       Exit Sub
    End If
    lbNuevo = True
    If oRsActividades.RecordCount > 0 Then
       oRsActividades.MoveFirst
       Do While Not oRsActividades.EOF
          If oRsActividades!Grupo = Val(Me.txtGrupo.Text) And oRsActividades!SubGrupo = Val(Me.txtSubGrupo.Text) Then
             lbNuevo = False
             Exit Do
          End If
          oRsActividades.MoveNext
       Loop
    End If
    If lbNuevo = True Then
        oRsActividades.AddNew
        oRsActividades!UPS = ml_ups
        oRsActividades!Grupo = Val(Me.txtGrupo.Text)
        oRsActividades!SubGrupo = Val(Me.txtSubGrupo.Text)
        oRsActividades!subgrupoOrden = Val(Me.txtSubGrupo.Text)
    End If
    oRsActividades!lab = Me.txtLab.Text
    oRsActividades!cpt_dx = Me.lblDxCpt.Tag
    oRsActividades!idTipo = IIf(Me.optCPT.Value = True, 1, 3)
    oRsActividades!EdadInicio = Val(Me.txtEdad.Text)
    oRsActividades!EdadFinal = Val(Me.txtEdadFin.Text)
    oRsActividades!idTipoEdad = Me.cmdTipoEdad.ListIndex + 1
    oRsActividades!PesoKgMenor = Val(Me.txtPesoInicial.Text)
    oRsActividades!PesoKgMayor = Val(Me.txtPesoFinal.Text)
    oRsActividades.Update
    
    mo_Formulario.HabilitarDeshabilitar FraAgregar(0), False
    LimpiaDatos
End Sub

Private Sub cmdNew_Click()
'   oRsActividades.AddNew
    mo_Formulario.HabilitarDeshabilitar FraAgregar(0), True
    LimpiaDatos
    txtGrupo.SetFocus
End Sub

Sub LimpiaDatos()
   txtGrupo.Text = ""
   txtSubGrupo.Text = ""
   txtLab.Text = ""
   txtEdad.Text = "0"
   txtEdadFin.Text = "140"
   cmdTipoEdad.ListIndex = 0
   txtPesoInicial.Text = "0"
   txtPesoFinal.Text = "200"
   lblDxCpt.Caption = ""
   lblDxCpt.Tag = ""
End Sub

Private Sub cmdSalir_Click()
     Me.Visible = False
End Sub



Private Sub cmdTipoEdad_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmdTipoEdad
End Sub



Private Sub Form_Load()
    mo_Formulario.HabilitarDeshabilitar FraAgregar(0), False
    ml_idCuentaAtencion = 0
    ml_NroFua = 0
    Me.Width = 11955
    Select Case ml_formLlamante
    Case "CPT"
         Set grdBusqueda.DataSource = mo_ReglasAdmision.ServiciosAtenSimultaneaMovXcorrelativo(ml_idCorrelativo, True)
         Me.Caption = "Consultorios que registran información al mismo tiempo"
         grdBusqueda.Caption = "Lista de Consultorios"
         grdBusqueda.Height = 2085
         grdBusqueda.Width = 7005
         Me.fraAceptar.Top = 2190
         Me.fraAceptar.Width = 7000
         Me.Height = 3795
         Me.Width = 7400
         Me.btnAceptar.Left = 2077
         Me.btnCancelar.Left = 3622
    Case "FUA"
         Set grdBusqueda.DataSource = mrs_oRsFua
         grdBusqueda.Caption = "Lista de FUAs"
         grdBusqueda.Height = 2085
         grdBusqueda.Width = 7005
         Me.fraAceptar.Top = 2190
         Me.fraAceptar.Width = 7000
         Me.Height = 3795
         Me.Width = 7400
         Me.btnAceptar.Left = 2077
         Me.btnCancelar.Left = 3622
    Case "ACTIVIDADES"
         If oRsItemsMasivosElegidos.State = 1 Then Set oRsItemsMasivosElegidos = Nothing
         With oRsItemsMasivosElegidos
              .Fields.Append "Grupo", adInteger
              .Fields.Append "SubGrupo", adInteger
              .Fields.Append "lab", adVarChar, 3, adFldIsNullable
              .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
              .Fields.Append "id", adVarChar, 20, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 255, adFldIsNullable
              .Fields.Append "Elija", adBoolean
              .Fields.Append "ElijaTipo", adInteger
              .Fields.Append "ElijaUPS", adInteger
              .Fields.Append "ElijaLab", adVarChar, 3, adFldIsNullable
              .Fields.Append "IdCuentaAtencion", adInteger, 4
              .Fields.Append "IdOrden", adInteger, 4
              .Fields.Append "Fua", adInteger
              .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
              .Fields.Append "IdServicio", adInteger
              .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
              .Fields.Append "idTipo", adInteger
              .Fields.Append "idServicioPaciente", adInteger
              .CursorType = adOpenKeyset
              .LockType = adLockOptimistic
              .Open
         End With
         mrs_oRsFua.MoveFirst
         Set grdBusqueda.DataSource = mrs_oRsFua
    
        Dim oRsLab As New Recordset
        Set oRsLab = mo_AdminServiciosComunes.DevuelveHIS_SITUACIOporDescripcion()
        With grdBusqueda.ValueLists.Add("LabLista").ValueListItems
           oRsLab.MoveFirst
           Do While Not oRsLab.EOF
              .Add Trim(oRsLab!valores), oRsLab!valores
              oRsLab.MoveNext
           Loop
        End With
        oRsLab.Close
        Set oRsLab = Nothing
        grdBusqueda.Bands(0).Columns("ElijaLab").ValueList = "LabLista"
        grdBusqueda.Bands(0).Columns("ElijaLab").ButtonDisplayStyle = ssButtonDisplayStyleAlways
         
         
         'mrs_oRsFua.MoveFirst
         'Set grdBusqueda.DataSource = mrs_oRsFua
         Dim lnFor As Integer
         With grdBusqueda.ValueLists.Add("ElijaTipoList").ValueListItems
             oRsTipoDx.MoveFirst
             Do While Not oRsTipoDx.EOF
                lnFor = oRsTipoDx!IdSubclasificacionDx - 100
                .Add lnFor, oRsTipoDx!DescripcionLarga  'oRsTipoDx!IdSubclasificacionDx, oRsTipoDx!DescripcionLarga
                oRsTipoDx.MoveNext
             Loop
         End With
         grdBusqueda.Bands(0).Columns("ElijaTipo").ValueList = "ElijaTipoList"
         grdBusqueda.Bands(0).Columns("ElijaTipo").ButtonDisplayStyle = ssButtonDisplayStyleAlways
         
         grdBusqueda.Bands(0).Columns("ElijaUPS").Hidden = True
         grdBusqueda.Bands(0).Columns("IdCuentaAtencion").Hidden = True
         grdBusqueda.Bands(0).Columns("IdOrden").Hidden = True
         grdBusqueda.Bands(0).Columns("Fua").Hidden = True
         grdBusqueda.Bands(0).Columns("Consultorio").Hidden = True
         grdBusqueda.Bands(0).Columns("IdServicio").Hidden = True
         grdBusqueda.Bands(0).Columns("FuaCodigoPrestacion").Hidden = True
         grdBusqueda.Bands(0).Columns("idTipo").Hidden = True
         grdBusqueda.Bands(0).Columns("idServicioPaciente").Hidden = True
         grdBusqueda.Bands(0).Columns("Elija").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("Elija").Header.Appearance.Font.Bold = True
         grdBusqueda.Bands(0).Columns("ElijaTipo").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("ElijaTipo").Header.Appearance.Font.Bold = True
         grdBusqueda.Bands(0).Columns("ElijaLab").Header.Appearance.BackColor = vbRed
         grdBusqueda.Bands(0).Columns("ElijaLab").Header.Appearance.Font.Bold = True
         '
         grdBusqueda.Caption = "Lista de ACTIVIDADES según EDAD,PESO,UPS"
         Me.Caption = "ACTIVIDADES"

    Case "SERVICIOS"
         FraAgregar(0).Visible = True
         Me.Caption = "Mantenimiento de ACTIVIDADES"
         grdBusqueda.Visible = False
         fraAceptar.Visible = False
         Me.Width = 15570
         lcSql = "select * from ServiciosAtenSimultaneaImpHIS where ups='" & ml_ups & "' order by grupo, subgrupoorden"
         oRsActividades.Open lcSql, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
         Set grdBusqueda.DataSource = oRsActividades
         grdBusqueda.Visible = True
         FraConsideraciones.Visible = True
         cmbTipoEd11.ListIndex = 0
    End Select
End Sub




Private Sub grdBusqueda_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    Select Case ml_formLlamante
    Case "ACTIVIDADES"
        If Cell.Column.key = "Elija" Then
           If Cell.Row.Cells("Elija").Value = True Then
              If Trim(Cell.Row.Cells("Lab").Value) <> "" Then
                 Cell.Row.Cells("ElijaLab").Value = Cell.Row.Cells("Lab").Value
              End If
           Else
              Cell.Row.Cells("ElijaLab").Value = ""
           End If
        End If
    End Select
End Sub

Private Sub grdBusqueda_DblClick()
    Dim rsRecordset As Recordset
    Set rsRecordset = grdBusqueda.DataSource
    Select Case ml_formLlamante
    Case "CPT"
         ml_idCuentaAtencion = rsRecordset("IdCuentaAtencion")
    Case "FUA"
         ml_NroFua = rsRecordset("fua")
         ml_FuaIdCuentaAtencion = IIf(IsNull(rsRecordset("idFuaIdCuentaAtencion")), 0, rsRecordset("idFuaIdCuentaAtencion"))
    Case "ACTIVIDADES"
         ml_idCuentaAtencion = 1
    End Select
    Me.Visible = False
End Sub

Private Sub grdBusqueda_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Select Case ml_formLlamante
    Case "CPT"
        grdBusqueda.Bands(0).Columns("IdServicioIngreso").Hidden = True
        grdBusqueda.Bands(0).Columns("idAtencion").Hidden = True
        grdBusqueda.Bands(0).Columns("IdCuentaAtencion").Width = 1000
        grdBusqueda.Bands(0).Columns("consultorio").Width = 3950
        grdBusqueda.Bands(0).Columns("consultorio").Header.Caption = "UPS"
    Case "FUA"
        grdBusqueda.Bands(0).Columns("codigoCIEsinPto").Hidden = True
        grdBusqueda.Bands(0).Columns("idFuaIdCuentaAtencion").Hidden = True
        grdBusqueda.Bands(0).Columns("fua").Width = 1000
        grdBusqueda.Bands(0).Columns("diagnosticos").Header.Caption = "Diagnósticos"
        grdBusqueda.Bands(0).Columns("diagnosticos").Width = 3950
        grdBusqueda.Bands(0).Columns("fuaCodigoPrestacion").Width = 1300
        grdBusqueda.Bands(0).Columns("fuaCodigoPrestacion").Header.Caption = "Cod.Prestación"
    Case "ACTIVIDADES"
        grdBusqueda.Bands(0).Columns("Grupo").Hidden = True
        grdBusqueda.Bands(0).Columns("subGrupo").Hidden = True
        grdBusqueda.Bands(0).Columns("grupoTIT").Width = 500
        grdBusqueda.Bands(0).Columns("grupoTIT").Header.Caption = "Grupo"
        grdBusqueda.Bands(0).Columns("lab").Width = 500
        grdBusqueda.Bands(0).Columns("tipo").Width = 500
        grdBusqueda.Bands(0).Columns("id").Width = 800
        grdBusqueda.Bands(0).Columns("nombre").Width = 4000
        
        
    Case "SERVICIOS"
        grdBusqueda.Bands(0).Columns("UPS").Hidden = True
        grdBusqueda.Bands(0).Columns("Grupo").Width = 500
        grdBusqueda.Bands(0).Columns("SubGrupo").Hidden = True
        grdBusqueda.Bands(0).Columns("subGrupoOrden").Width = 1000
        grdBusqueda.Bands(0).Columns("lab").Width = 500
        grdBusqueda.Bands(0).Columns("cpt_dx").Width = 1000
        grdBusqueda.Bands(0).Columns("idTipo").Hidden = True
        grdBusqueda.Bands(0).Columns("EdadInicio").Width = 1500
        grdBusqueda.Bands(0).Columns("EdadFinal").Width = 1500
        grdBusqueda.Bands(0).Columns("idTipoEdad").Width = 2000
        grdBusqueda.Bands(0).Columns("idTipoEdad").Header.Caption = "1(años),2(meses),3(días) "
        grdBusqueda.Bands(0).Columns("PesoKgMenor").Width = 1500
        grdBusqueda.Bands(0).Columns("PesoKgMayor").Width = 1500
    End Select
    mo_Apariencia.ConfigurarFilasBiColores grdBusqueda, sighentidades.GrillaConFilasBicolor
End Sub

Private Sub grdBusqueda_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
   If KeyAscii = 13 Then
      grdBusqueda_DblClick
   End If
End Sub

Property Get ItemsMasivosElegidos() As Recordset
    If lbPulsoBotonAceptar = True Then
        Dim oRsTmp99 As New Recordset
        Set oRsTmp99 = grdBusqueda.DataSource
        oRsTmp99.MoveFirst
        Do While Not oRsTmp99.EOF
            If oRsTmp99!elija = True Then
                oRsItemsMasivosElegidos.AddNew
                oRsItemsMasivosElegidos!Grupo = oRsTmp99!Grupo
                oRsItemsMasivosElegidos!SubGrupo = oRsTmp99!SubGrupo
                oRsItemsMasivosElegidos!lab = oRsTmp99!lab
                oRsItemsMasivosElegidos!ID = oRsTmp99!ID
                oRsItemsMasivosElegidos!tipo = oRsTmp99!tipo
                oRsItemsMasivosElegidos!nombre = oRsTmp99!nombre
                oRsItemsMasivosElegidos!elija = oRsTmp99!elija
                oRsItemsMasivosElegidos!elijaTipo = oRsTmp99!elijaTipo + 100
                oRsItemsMasivosElegidos!ElijaUPS = oRsTmp99!ElijaUPS
                oRsItemsMasivosElegidos!ElijaLab = oRsTmp99!ElijaLab
                
                oRsItemsMasivosElegidos!idOrden = oRsTmp99!idOrden
                oRsItemsMasivosElegidos!fua = oRsTmp99!fua
                oRsItemsMasivosElegidos!Consultorio = oRsTmp99!Consultorio
                oRsItemsMasivosElegidos!idServicio = oRsTmp99!idServicio
                oRsItemsMasivosElegidos!FuaCodigoPrestacion = oRsTmp99!FuaCodigoPrestacion
                oRsItemsMasivosElegidos!idTipo = oRsTmp99!idTipo
                oRsItemsMasivosElegidos!idCuentaAtencion = oRsTmp99!idCuentaAtencion
                oRsItemsMasivosElegidos!idServicioPaciente = oRsTmp99!idServicioPaciente
            End If
            oRsTmp99.MoveNext
        Loop
        Set oRsTmp99 = Nothing
    
'        Dim oRow As SSRow
'        grdBusqueda.Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
'        grdBusqueda.Bands(0).Columns("Elija").SortIndicator = ssSortIndicatorAscending
'        Set oRow = Me.grdBusqueda.GetRow(ssChildRowFirst)
'        If Not oRow Is Nothing Then
'        oRow.GetSibling (ssSiblingRowFirst)
'                Do While oRow.HasNextSibling
'                    Set oRow = oRow.GetSibling(ssSiblingRowNext)
'                    If oRow.Cells("Elija").Value = True Then
'                        oRsItemsMasivosElegidos.AddNew
'                        oRsItemsMasivosElegidos!Grupo = oRow.Cells("Grupo").Value
'                        oRsItemsMasivosElegidos!SubGrupo = oRow.Cells("SubGrupo").Value
'                        oRsItemsMasivosElegidos!lab = oRow.Cells("lab").Value
'                        oRsItemsMasivosElegidos!Id = oRow.Cells("id").Value
'                        oRsItemsMasivosElegidos!tipo = oRow.Cells("tipo").Value
'                        oRsItemsMasivosElegidos!nombre = oRow.Cells("nombre").Value
'                        oRsItemsMasivosElegidos!elija = oRow.Cells("elija").Value
'                        oRsItemsMasivosElegidos!elijaTipo = oRow.Cells("elijaTipo").Value + 100
'                        oRsItemsMasivosElegidos!ElijaUPS = oRow.Cells("ElijaUPS").Value
'                        oRsItemsMasivosElegidos!ElijaLab = oRow.Cells("ElijaLab").Value
'
'                        oRsItemsMasivosElegidos!IdOrden = oRow.Cells("idOrden").Value
'                        oRsItemsMasivosElegidos!Fua = oRow.Cells("Fua").Value
'                        oRsItemsMasivosElegidos!Consultorio = oRow.Cells("Consultorio").Value
'                        oRsItemsMasivosElegidos!IdServicio = oRow.Cells("idServicio").Value
'                        oRsItemsMasivosElegidos!FuaCodigoPrestacion = oRow.Cells("FuaCodigoPrestacion").Value
'                        oRsItemsMasivosElegidos!idTipo = oRow.Cells("idTipo").Value
'                        oRsItemsMasivosElegidos!idCuentaAtencion = oRow.Cells("idCuentaAtencion").Value
'                        oRsItemsMasivosElegidos!idServicioPaciente = oRow.Cells("idServicioPaciente").Value
'                    End If
'                Loop
'
'        End If
'
'        Set oRow = Nothing
       
        Set ItemsMasivosElegidos = oRsItemsMasivosElegidos.Clone
    End If
End Sub





Private Sub optCPT_Click(Value As Integer)
    Me.lblDxCpt.Caption = ""
    Me.lblDxCpt.Tag = ""

End Sub

Private Sub optDx_Click(Value As Integer)
    Me.lblDxCpt.Caption = ""
    Me.lblDxCpt.Tag = ""

End Sub

Private Sub txtEdad_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtEdad
End Sub



Private Sub txtEdad_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtEdadFin_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtEdadFin
End Sub

Private Sub txtEdadFin_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
        mo_Teclado.RealizarNavegacion KeyCode, txtGrupo

End Sub





Private Sub txtGrupo_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

Private Sub txtLab_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtLab
End Sub





Private Sub txtPesoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtPesoFinal
End Sub

Private Sub txtPesoFinal_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtPesoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtPesoInicial
End Sub

Private Sub txtPesoInicial_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub

Private Sub txtSubGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
mo_Teclado.RealizarNavegacion KeyCode, txtSubGrupo
End Sub

Private Sub txtSubGrupo_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If

End Sub
