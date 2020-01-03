VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form SolicitudHistoriasReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de solicitud de historias clínicas"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "SolicitudHistoriasReporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBuscarRespArchivo 
      Caption         =   "..."
      Height          =   315
      Left            =   3030
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   240
      Width           =   315
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   15
      Top             =   3630
      Width           =   7620
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SolicitudHistoriasReporte.frx":0CCA
         DownPicture     =   "SolicitudHistoriasReporte.frx":118E
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
         Left            =   3930
         Picture         =   "SolicitudHistoriasReporte.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SolicitudHistoriasReporte.frx":1B66
         DownPicture     =   "SolicitudHistoriasReporte.frx":1FC6
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
         Left            =   2400
         Picture         =   "SolicitudHistoriasReporte.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3570
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   7635
      Begin VB.CheckBox chkHaySaltoDePagXconsultorio 
         Caption         =   "Salto de página x Consultorio"
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
         Left            =   1950
         TabIndex        =   26
         Top             =   2640
         Value           =   1  'Checked
         Width           =   5475
      End
      Begin VB.CheckBox chkCitasPagadas 
         Caption         =   "Incluir solo CITAS PAGADAS (si tiene plan=Particular)"
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
         Left            =   1950
         TabIndex        =   25
         Top             =   2340
         Width           =   5475
      End
      Begin VB.CheckBox chkIncluyeHS 
         Caption         =   "Incluir Historias q´salieron del ARCHIVO CLINICO a SERVICIOS"
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
         Left            =   1950
         TabIndex        =   24
         Top             =   2040
         Value           =   1  'Checked
         Width           =   5475
      End
      Begin VB.CheckBox chkHistoricos 
         Caption         =   "Incluir atenciones históricas"
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
         Left            =   1950
         TabIndex        =   21
         Top             =   1725
         Width           =   5115
      End
      Begin VB.ComboBox cmbIdTipoServicio 
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
         Left            =   1950
         TabIndex        =   19
         Top             =   1305
         Width           =   5145
      End
      Begin VB.TextBox txtNombreEmpleado 
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
         Left            =   3360
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Width           =   3690
      End
      Begin VB.TextBox txtIdEmpleado 
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
         Left            =   1950
         TabIndex        =   0
         Top             =   210
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFechaRequeridaDesde 
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Top             =   585
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
      Begin MSMask.MaskEdBox txtFechaRequeridaHasta 
         Height          =   315
         Left            =   4890
         TabIndex        =   2
         Top             =   600
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
      Begin MSMask.MaskEdBox txtFechaSolicitudDesde 
         Height          =   315
         Left            =   1950
         TabIndex        =   3
         Top             =   960
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
      Begin MSMask.MaskEdBox txtFechaSolicitudHasta 
         Height          =   315
         Left            =   3765
         TabIndex        =   4
         Top             =   975
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
      Begin GalenHos.XP_ProgressBar progressRpt 
         Height          =   300
         Left            =   1920
         TabIndex        =   17
         Top             =   3120
         Width           =   5145
         _extentx        =   9075
         _extenty        =   529
         font            =   "SolicitudHistoriasReporte.frx":28B0
         brushstyle      =   0
         color           =   6956042
      End
      Begin MSMask.MaskEdBox txtHoraReqIni 
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraReqFin 
         Height          =   315
         Left            =   6300
         TabIndex        =   23
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblIdTipoServicio 
         Caption         =   "Tipo de servicio"
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
         TabIndex        =   20
         Top             =   1350
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Al"
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
         Left            =   3480
         TabIndex        =   14
         Top             =   990
         Width           =   150
      End
      Begin VB.Label Label5 
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
         Height          =   210
         Left            =   1620
         TabIndex        =   13
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha solicitud"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Resp. de archivo"
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
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Al"
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
         Left            =   4410
         TabIndex        =   10
         Top             =   660
         Width           =   150
      End
      Begin VB.Label Label2 
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
         Height          =   210
         Left            =   1620
         TabIndex        =   9
         Top             =   645
         Width           =   345
      End
      Begin VB.Label lblFechaRequerida 
         Caption         =   "Fecha Requerida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   645
         Width           =   1350
      End
   End
End
Attribute VB_Name = "SolicitudHistoriasReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************daniel barrantes**************
'***************se incluyó filtro por HORA DE CITA
'***************
Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_AdminReglasCOmunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_cmbIdTipoServicio As New SIGHEntidades.ListaDespleglable
Dim ms_TipoReporte
Dim ml_idUsuario As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let TipoReporte(sValue As String)
    ms_TipoReporte = sValue
End Property


Private Sub btnAceptar_Click()
Dim oRptSolicitud As New RptSolicitudHistoria

    
    If mo_cmbIdTipoServicio.BoundText = "" Then
        MsgBox "Ingrese el tipo de servicio", vbInformation, Me.Caption
        Exit Sub
    End If

    oRptSolicitud.IdEmpleado = Val(Me.txtIdEmpleado.Tag)
    If Me.txtFechaRequeridaDesde = SIGHEntidades.FECHA_VACIA_DMY Then
       oRptSolicitud.FechaRequeridaDesde = 0
    Else
       oRptSolicitud.FechaRequeridaDesde = CDate(Format(Me.txtFechaRequeridaDesde & " " & Me.txtHoraReqIni.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    End If
    If Me.txtFechaRequeridaHasta = SIGHEntidades.FECHA_VACIA_DMY Then
       oRptSolicitud.FechaRequeridaHasta = 0
    Else
       oRptSolicitud.FechaRequeridaHasta = CDate(Format(Me.txtFechaRequeridaHasta & " " & Me.txtHoraReqFin.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    End If
    If Me.txtFechaSolicitudDesde = SIGHEntidades.FECHA_VACIA_DMY Then
       oRptSolicitud.FechaSolicitudDesde = 0
    Else
       oRptSolicitud.FechaSolicitudDesde = CDate(Format(Me.txtFechaSolicitudDesde & " 00:00:01", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    End If
    If Me.txtFechaSolicitudHasta = SIGHEntidades.FECHA_VACIA_DMY Then
       oRptSolicitud.FechaSolicitudHasta = 0
    Else
       oRptSolicitud.FechaSolicitudHasta = CDate(Format(Me.txtFechaSolicitudHasta & " 23:59:59", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    End If
    oRptSolicitud.HoraReqIni = Me.txtHoraReqIni.Text
    oRptSolicitud.HoraReqFin = Me.txtHoraReqFin.Text
    oRptSolicitud.Historicos = IIf(chkHistoricos.Value = 1, True, False)
    oRptSolicitud.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oRptSolicitud.IncluyeHistoriasQueSalieron = chkIncluyeHS.Value
    oRptSolicitud.SoloCitasPagadas = IIf(chkCitasPagadas.Value = 1, True, False)
    Set oRptSolicitud.progressRpt = Me.progressRpt
    
    Select Case ms_TipoReporte
    Case "RPT_HISTORIAS_SERVICIO"
        oRptSolicitud.CrearReporteHistoriaSolicitadas
    Case "RPT_HISTORIAS_MEDICO"
        If chkHaySaltoDePagXconsultorio.Value = 0 Then
           oRptSolicitud.CrearReporteHistoriaSolicitadasDeCEPorMedico
        Else
            Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.DestinoReporte = sghPantalla
            oRptClaseCry.idUsuario = Val(Me.txtIdEmpleado.Tag)
            If Me.txtFechaRequeridaDesde = SIGHEntidades.FECHA_VACIA_DMY Then
               oRptClaseCry.FechaInicio = 0
            Else
               oRptClaseCry.FechaInicio = CDate(Format(Me.txtFechaRequeridaDesde & " " & Me.txtHoraReqIni.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            End If
            If Me.txtFechaRequeridaHasta = SIGHEntidades.FECHA_VACIA_DMY Then
               oRptClaseCry.FechaFin = 0
            Else
               oRptClaseCry.FechaFin = CDate(Format(Me.txtFechaRequeridaHasta & " " & Me.txtHoraReqFin.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            End If
            If Me.txtFechaSolicitudDesde = SIGHEntidades.FECHA_VACIA_DMY Then
               oRptClaseCry.FechaSolicitudDesde = 0
            Else
               oRptClaseCry.FechaSolicitudDesde = CDate(Format(Me.txtFechaSolicitudDesde & " 00:00:01", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            End If
            If Me.txtFechaSolicitudHasta = SIGHEntidades.FECHA_VACIA_DMY Then
               oRptClaseCry.FechaSolicitudHasta = 0
            Else
               oRptClaseCry.FechaSolicitudHasta = CDate(Format(Me.txtFechaSolicitudHasta & " 23:59:59", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            End If
            oRptClaseCry.lcTipoServicio = mo_cmbIdTipoServicio.BoundText
            oRptClaseCry.IncluyeHistoriasQueSalieron = chkIncluyeHS.Value
            oRptClaseCry.TipoReporte = "HcXmedicoXpagina"
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
            Me.MousePointer = 1
        End If
    End Select


End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
       mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub Form_Initialize()
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
End Sub

Private Sub Form_Load()

   Me.txtFechaRequeridaDesde.Text = Date
   Me.txtHoraReqIni.Text = "00:00": Me.txtHoraReqFin.Text = "23:59"
   Me.txtFechaRequeridaHasta.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
   '
   mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
   mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
   Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
   mo_cmbIdTipoServicio.BoundText = 1
   '
   Dim oDOEmpleado As New dOEmpleado
   Set oDOEmpleado = mo_AdminReglasCOmunes.EmpleadosSeleccionarPorId(ml_idUsuario)
   If Not oDOEmpleado Is Nothing Then
        txtIdEmpleado.Tag = oDOEmpleado.IdEmpleado
        txtIdEmpleado.Text = oDOEmpleado.CodigoPlanilla
        txtNombreEmpleado = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
   End If
   Set oDOEmpleado = Nothing
End Sub

Private Sub txtFechaRequeridaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaRequeridaDesde
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaRequeridaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaRequeridaHasta
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaSolicitudDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitudDesde
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaSolicitudHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitudHasta
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdEmpleado
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdEmpleado_LostFocus()
    CompletarDatosDeEmpleadoEnElLostFocus txtIdEmpleado, Me.txtNombreEmpleado
    mo_Formulario.MarcarComoVacio txtIdEmpleado
End Sub

Private Sub txtIdEmpleado_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub btnBuscarRespArchivo_Click()
    CompletarDatosResponsable Me.txtIdEmpleado, Me.txtNombreEmpleado
End Sub
Sub CompletarDatosResponsable(txtIdResponsable As TextBox, txtNombreResponsable As TextBox)
'Dim oBusqueda As New EmpleadosBusqueda
Dim oBusqueda As New SIGHCatalogos.clEmpleadosBusqueda
Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    'oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminReglasCOmunes.EmpleadosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtIdResponsable.Tag = oDOEmpleado.IdEmpleado
            txtIdResponsable.Text = oDOEmpleado.CodigoPlanilla
            txtNombreResponsable = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        End If
    End If

End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Sub CompletarDatosDeEmpleadoEnElLostFocus(txtCodigoPlanilla As TextBox, txtNombre As TextBox)
Dim oDOEmpleado As New dOEmpleado

        If mo_AdminReglasCOmunes.EmpleadosSeleccionarPorCodigo(txtCodigoPlanilla.Text, oDOEmpleado) Then
            txtCodigoPlanilla.Tag = oDOEmpleado.IdEmpleado
            txtCodigoPlanilla.Text = oDOEmpleado.CodigoPlanilla
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.ApellidoMaterno + " " + oDOEmpleado.Nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub

