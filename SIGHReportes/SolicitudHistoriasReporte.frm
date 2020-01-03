VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form SolicitudHistoriasReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de solicitud de historias clínicas"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   Icon            =   "SolicitudHistoriasReporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SIGHReportes.ucElegirTurno ucElegirTurno1 
      Height          =   1890
      Left            =   7710
      TabIndex        =   32
      Top             =   15
      Width           =   2085
      _extentx        =   3678
      _extenty        =   3334
   End
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
      Top             =   5280
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
      Height          =   5130
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   7635
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   1800
         TabIndex        =   28
         Top             =   3240
         Width           =   5655
         Begin Threed.SSOption optListaXconsultorios 
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
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
            Caption         =   "Listado ordenado por Médicos"
         End
         Begin Threed.SSOption optConsultorioXpagina 
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   5295
            _ExtentX        =   9340
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
            Caption         =   "Lista por cada Médico (salto de página x Médico)"
         End
         Begin Threed.SSOption optConsultorioConFF 
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   5295
            _ExtentX        =   9340
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
            Caption         =   "Listado por Médico incluyendo Fuente Financiamiento"
            Value           =   -1
         End
      End
      Begin VB.TextBox txtUltimoDigitoHC 
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
         Left            =   2025
         TabIndex        =   26
         Top             =   1695
         Width           =   1395
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
         Left            =   2025
         TabIndex        =   25
         Top             =   2700
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
         Left            =   2025
         TabIndex        =   24
         Top             =   2400
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
         Left            =   2025
         TabIndex        =   21
         Top             =   2085
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
         Left            =   2025
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
         Left            =   3435
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
         Left            =   2025
         TabIndex        =   0
         Top             =   210
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFechaRequeridaDesde 
         Height          =   315
         Left            =   2025
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
         Left            =   2025
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
         Left            =   3840
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
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   1800
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   17
         Top             =   4680
         Visible         =   0   'False
         Width           =   5655
      End
      Begin MSMask.MaskEdBox txtHoraReqIni 
         Height          =   315
         Left            =   3435
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
      Begin VB.Label Label7 
         Caption         =   "Ultimos digitos de la HC"
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
         TabIndex        =   27
         Top             =   1710
         Width           =   1905
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
         Left            =   120
         TabIndex        =   20
         Top             =   1320
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
         Left            =   3555
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
         Left            =   120
         TabIndex        =   12
         Top             =   990
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
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
         Left            =   120
         TabIndex        =   8
         Top             =   615
         Width           =   1350
      End
   End
End
Attribute VB_Name = "SolicitudHistoriasReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Solicitud de Historias
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
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
'Dim oRptSolicitud As New RptSolicitudHistoria
Dim oRptSolicitud As New clSolicitudHistorias

    
    If mo_cmbIdTipoServicio.BoundText = "" Then
        MsgBox "Ingrese el tipo de servicio", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Me.txtFechaRequeridaDesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(Me.txtFechaRequeridaDesde, "DD/MM/AAAA") Then
            MsgBox "La fecha requerida inicial, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
        If Me.txtHoraReqIni = SIGHEntidades.HORA_VACIA_HM Then
            MsgBox "Ingrese la hora requerida inicial", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not SIGHEntidades.EsHora(txtHoraReqIni) Then
                MsgBox "La hora requerida inicial, no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
    End If
    
    If Me.txtFechaRequeridaHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(Me.txtFechaRequeridaHasta, "DD/MM/AAAA") Then
            MsgBox "La fecha requerida final, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
        If Me.txtHoraReqFin = SIGHEntidades.HORA_VACIA_HM Then
            MsgBox "Ingrese la hora requerida final", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not SIGHEntidades.EsHora(txtHoraReqFin) Then
                MsgBox "La hora requerida final, no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    If CDate(Me.txtFechaRequeridaDesde.Text & " " & Me.txtHoraReqIni.Text) > CDate(Me.txtFechaRequeridaHasta.Text & " " & Me.txtHoraReqFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Sub
    End If

    If Me.txtFechaSolicitudDesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(Me.txtFechaSolicitudDesde, "DD/MM/AAAA") Then
            MsgBox "La fecha de solicitud final, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
        
    End If

    If Me.txtFechaSolicitudHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(Me.txtFechaSolicitudHasta, "DD/MM/AAAA") Then
            MsgBox "La fecha de solicitud final, no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
        If CDate(Me.txtFechaSolicitudDesde.Text) > CDate(Me.txtFechaSolicitudHasta.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Sub
        End If
    End If
    

    oRptSolicitud.IdEmpleado = Val(Me.txtIdEmpleado.Tag)
    If Me.txtFechaRequeridaDesde = SIGHEntidades.FECHA_VACIA_DMY Then
       oRptSolicitud.FechaRequeridaDesde = 0
    Else
       oRptSolicitud.FechaRequeridaDesde = CDate(Format(Me.txtFechaRequeridaDesde, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
    End If
    If Me.txtFechaRequeridaHasta = SIGHEntidades.FECHA_VACIA_DMY Then
       oRptSolicitud.FechaRequeridaHasta = 0
    Else
       oRptSolicitud.FechaRequeridaHasta = CDate(Format(Me.txtFechaRequeridaHasta, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
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
    oRptSolicitud.UltimosDigitosHC = Trim(txtUltimoDigitoHC.Text)
    
    Select Case ms_TipoReporte
    Case "RPT_HISTORIAS_SERVICIO"
        oRptSolicitud.CrearReporteHistoriaSolicitadas Me.hwnd
    Case "RPT_HISTORIAS_MEDICO"
    
        If optListaXconsultorios.Value = True Then
           oRptSolicitud.CrearReporteHistoriaSolicitadasDeCEPorMedico Me.hwnd
        Else
            Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.DestinoReporte = sghPantalla
            oRptClaseCry.idUsuario = Val(Me.txtIdEmpleado.Tag)
            If Me.txtFechaRequeridaDesde = SIGHEntidades.FECHA_VACIA_DMY Then
               oRptClaseCry.FechaInicio = 0
            Else
               oRptClaseCry.FechaInicio = CDate(Format(Me.txtFechaRequeridaDesde & " " & Me.txtHoraReqIni.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
               oRptClaseCry.HoraInicio = Me.txtHoraReqIni.Text
            End If
            If Me.txtFechaRequeridaHasta = SIGHEntidades.FECHA_VACIA_DMY Then
               oRptClaseCry.FechaFin = 0
            Else
               oRptClaseCry.FechaFin = CDate(Format(Me.txtFechaRequeridaHasta & " " & Me.txtHoraReqFin.Text, SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
               oRptClaseCry.HoraFin = Me.txtHoraReqFin.Text
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
            If optConsultorioXpagina.Value = True Then
               oRptClaseCry.TipoReporte = "HcXmedicoXpagina"
            Else
               oRptClaseCry.TipoReporte = "HcXmedico"
            End If
            oRptClaseCry.UltimosDigitosHC = Trim(txtUltimoDigitoHC.Text)
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
   ucElegirTurno1.Inicializar
   
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
        txtNombreEmpleado = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.apellidoMaterno + " " + oDOEmpleado.nombres
   End If
   Set oDOEmpleado = Nothing
   If ms_TipoReporte = "RPT_HISTORIAS_SERVICIO" Then
      Frame2.Visible = False
   End If
End Sub

Private Sub txtFechaRequeridaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaRequeridaDesde
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaRequeridaDesde_LostFocus()
    If txtFechaRequeridaDesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaRequeridaDesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaRequeridaDesde = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaRequeridaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaRequeridaHasta
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaRequeridaHasta_LostFocus()
    If txtFechaRequeridaHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaRequeridaHasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaRequeridaHasta = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaSolicitudDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitudDesde
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFechaSolicitudDesde_LostFocus()
    If txtFechaSolicitudDesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaSolicitudDesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaSolicitudDesde = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaSolicitudHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaSolicitudHasta
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaSolicitudHasta_LostFocus()
    If txtFechaSolicitudHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaSolicitudHasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaSolicitudHasta = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
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
Dim oBusqueda As New SIGHNegocios.BuscaEmpleados
Dim oDOEmpleado As New dOEmpleado
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOEmpleado = mo_AdminReglasCOmunes.EmpleadosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOEmpleado Is Nothing Then
            txtIdResponsable.Tag = oDOEmpleado.IdEmpleado
            txtIdResponsable.Text = oDOEmpleado.CodigoPlanilla
            txtNombreResponsable = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.apellidoMaterno + " " + oDOEmpleado.nombres
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
            txtNombre = oDOEmpleado.ApellidoPaterno + " " + oDOEmpleado.apellidoMaterno + " " + oDOEmpleado.nombres
        Else
            txtCodigoPlanilla.Tag = ""
            txtCodigoPlanilla = ""
            txtNombre = ""
        End If
End Sub

Private Sub txtHoraReqFin_LostFocus()
    If txtHoraReqFin <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHoraReqFin) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHoraReqFin = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub txtHoraReqIni_LostFocus()
    If txtHoraReqIni <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHoraReqIni) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHoraReqIni = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub ucElegirTurno1_SeModificoTurnos(lnTurno As SIGHEntidades.sghTurnos, lcHrMinicio As String, lcHrMfinal As String, lcHrTinicio As String, lcHrTfinal As String)
    Select Case lnTurno
    Case sghTurnoTarde
        txtHoraReqIni.Text = lcHrTinicio
        txtHoraReqFin.Text = lcHrTfinal
    Case sghTurnoManana
        txtHoraReqIni.Text = lcHrMinicio
        txtHoraReqFin.Text = lcHrMfinal
    Case Else
        txtHoraReqIni.Text = "00:00"
        txtHoraReqFin.Text = "23:59"
    End Select
End Sub
