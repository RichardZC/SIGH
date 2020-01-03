VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form CuposAsignadosReporte 
   Caption         =   "Reporte de Cupos Asignados "
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   Icon            =   "CuposAsignadosReporte.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2550
      Left            =   45
      TabIndex        =   12
      Top             =   60
      Width           =   7215
      Begin SIGHReportes.XP_ProgressBar progressRpt 
         Height          =   345
         Left            =   930
         TabIndex        =   19
         Top             =   1320
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   12937777
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
         Left            =   5745
         TabIndex        =   5
         Top             =   885
         Visible         =   0   'False
         Width           =   3915
      End
      Begin MSMask.MaskEdBox txtFechaRequeridaDesde 
         Height          =   315
         Left            =   1935
         TabIndex        =   1
         Top             =   540
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
         Left            =   3765
         TabIndex        =   2
         Top             =   540
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
         Left            =   1935
         TabIndex        =   3
         Top             =   915
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
         Top             =   915
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
      Begin MSMask.MaskEdBox txtFSolIni 
         Height          =   315
         Left            =   1890
         TabIndex        =   7
         Top             =   2070
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
      Begin MSMask.MaskEdBox txtFSolFin 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   2070
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
      Begin Threed.SSOption OptServ 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Sólo cupos ya asignados por Servicios"
      End
      Begin Threed.SSOption optMedicos 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1710
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cupos Programados/Asignados por Médicos"
         Value           =   -1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F. Requerida"
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
         Left            =   855
         TabIndex        =   18
         Top             =   2115
         Width           =   1020
      End
      Begin VB.Label lblFechaRequerida 
         AutoSize        =   -1  'True
         Caption         =   "F Requerida"
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
         Left            =   900
         TabIndex        =   17
         Top             =   570
         Width           =   960
      End
      Begin VB.Label Label3 
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
         Height          =   285
         Left            =   3480
         TabIndex        =   16
         Top             =   585
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "F. solicitud"
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
         Left            =   900
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Height          =   285
         Left            =   3465
         TabIndex        =   14
         Top             =   960
         Width           =   345
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
         Left            =   5745
         TabIndex        =   13
         Top             =   525
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   11
      Top             =   2655
      Width           =   7230
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CuposAsignadosReporte.frx":0CCA
         DownPicture     =   "CuposAsignadosReporte.frx":112A
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
         Left            =   2070
         Picture         =   "CuposAsignadosReporte.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CuposAsignadosReporte.frx":1A14
         DownPicture     =   "CuposAsignadosReporte.frx":1ED8
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
         Left            =   3600
         Picture         =   "CuposAsignadosReporte.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CuposAsignadosReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Cupos Asignados
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

Private Sub btnAceptar_Click()
Dim oRptSolicitud As New clCuposAsignadosRep
    
    If mo_cmbIdTipoServicio.BoundText = "" Then
        MsgBox "Ingrese el tipo de servicio", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If OptServ.Value Then
        If Me.txtFechaRequeridaDesde <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not SIGHEntidades.EsFecha(Me.txtFechaRequeridaDesde, "DD/MM/AAAA") Then
                MsgBox "La fecha requerida inicial no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFechaRequeridaHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not SIGHEntidades.EsFecha(Me.txtFechaRequeridaHasta, "DD/MM/AAAA") Then
                MsgBox "La fecha requerida final no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFechaSolicitudDesde <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not SIGHEntidades.EsFecha(Me.txtFechaSolicitudDesde, "DD/MM/AAAA") Then
                MsgBox "La fecha de solicitud inicial no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFechaSolicitudHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not SIGHEntidades.EsFecha(Me.txtFechaSolicitudHasta, "DD/MM/AAAA") Then
                MsgBox "La fecha de solicitud final no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If IsDate(Me.txtFechaSolicitudDesde.Text) And IsDate(Me.txtFechaSolicitudHasta.Text) Then
            If CDate(Me.txtFechaSolicitudDesde.Text) > CDate(Me.txtFechaSolicitudHasta.Text) Then
               MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
               Exit Sub
            End If
        End If
        Me.MousePointer = 11
        oRptSolicitud.FechaRequeridaDesde = IIf(Me.txtFechaRequeridaDesde = SIGHEntidades.FECHA_VACIA_DMY, 0, Me.txtFechaRequeridaDesde)
        oRptSolicitud.FechaRequeridaHasta = IIf(Me.txtFechaRequeridaHasta = SIGHEntidades.FECHA_VACIA_DMY, 0, Me.txtFechaRequeridaHasta)
        oRptSolicitud.FechaSolicitudDesde = IIf(Me.txtFechaSolicitudDesde = SIGHEntidades.FECHA_VACIA_DMY, 0, Me.txtFechaSolicitudDesde)
        oRptSolicitud.FechaSolicitudHasta = IIf(Me.txtFechaSolicitudHasta = SIGHEntidades.FECHA_VACIA_DMY, 0, Me.txtFechaSolicitudHasta)
        oRptSolicitud.idTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
        'Set oRptSolicitud.progressRpt = Me.progressRpt
        oRptSolicitud.CrearReporteCuposAsignados Me.hwnd
        Me.MousePointer = 1
    Else
        If Me.txtFSolIni <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not SIGHEntidades.EsFecha(Me.txtFSolIni, "DD/MM/AAAA") Then
                MsgBox "La fecha requerida inicial no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFSolFin <> SIGHEntidades.FECHA_VACIA_DMY Then
            If Not SIGHEntidades.EsFecha(Me.txtFSolFin, "DD/MM/AAAA") Then
                MsgBox "La fecha requerida final no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If CDate(Me.txtFSolIni.Text) > CDate(Me.txtFSolFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Sub
        End If
        Me.MousePointer = 11
        oRptSolicitud.FechaSolicitudDesde = IIf(Me.txtFSolIni.Text = SIGHEntidades.FECHA_VACIA_DMY, 0, Me.txtFSolIni.Text)
        oRptSolicitud.FechaSolicitudHasta = IIf(Me.txtFSolFin.Text = SIGHEntidades.FECHA_VACIA_DMY, 0, Me.txtFSolFin.Text)
        oRptSolicitud.CrearReporteCuposAsignadosVaciosPorMedico Me.hwnd
        Me.MousePointer = 1
    End If

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

    Me.txtFechaRequeridaDesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaRequeridaHasta.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    
    txtFSolIni.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    txtFSolFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)

   mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
   mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
   Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
   
   mo_cmbIdTipoServicio.BoundText = 1
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoServicio, False
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

Private Sub txtFechaSolicitudHasta_LostFocus()
    If txtFechaSolicitudHasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaSolicitudHasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaSolicitudHasta = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFSolFin_LostFocus()
    If txtFSolFin <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFSolFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFSolFin = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFSolIni_LostFocus()
    If txtFSolIni <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFSolIni, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFSolIni = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
