VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AHCEgresoMedico24 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historias Clínicas en condición Alta Médica que pasadas 48 hr no regresan al Archivo"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   45
      TabIndex        =   4
      Top             =   2025
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCEgresoMedico24.frx":0000
         DownPicture     =   "AHCEgresoMedico24.frx":0460
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
         Left            =   1320
         Picture         =   "AHCEgresoMedico24.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCEgresoMedico24.frx":0D4A
         DownPicture     =   "AHCEgresoMedico24.frx":120E
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
         Left            =   2850
         Picture         =   "AHCEgresoMedico24.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   5370
      Begin VB.ComboBox cmbTipoServicio 
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
         ItemData        =   "AHCEgresoMedico24.frx":1BE6
         Left            =   1680
         List            =   "AHCEgresoMedico24.frx":1BF0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   3555
      End
      Begin VB.ComboBox cmbOrden 
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
         ItemData        =   "AHCEgresoMedico24.frx":1C11
         Left            =   1680
         List            =   "AHCEgresoMedico24.frx":1C1B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1050
         Width           =   3555
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   620
         Width           =   3555
      End
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   135
         ScaleHeight     =   240
         ScaleWidth      =   5010
         TabIndex        =   2
         Top             =   2280
         Width           =   5070
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1470
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   3795
         TabIndex        =   12
         Top             =   1440
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label Label4 
         Caption         =   "F.Alta Médica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   14
         Top             =   1515
         Width           =   1110
      End
      Begin VB.Label Label3 
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
         Left            =   3300
         TabIndex        =   13
         Top             =   1485
         Width           =   435
      End
      Begin VB.Label Departamento 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Servicio"
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
         Left            =   135
         TabIndex        =   10
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Orden del Rep"
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
         Left            =   135
         TabIndex        =   7
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Responsable"
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
         Left            =   135
         TabIndex        =   3
         Top             =   660
         Width           =   1395
      End
   End
End
Attribute VB_Name = "AHCEgresoMedico24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Historias que no han regresado en 24 horas
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdResponsable As New SIGHEntidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptAHCEgresoMedico24 As New RptAHCEgresoMedico24
        oRptAHCEgresoMedico24.IdResponsable = Val(mo_cmbIdResponsable.BoundText)
        oRptAHCEgresoMedico24.OrdenFiltro = IIf(cmbOrden.ListIndex = 0, "HC", "Paciente")
        oRptAHCEgresoMedico24.FechaInicio = txtFechaInicio.Text
        oRptAHCEgresoMedico24.FechaFin = txtFechaFin.Text
        oRptAHCEgresoMedico24.TipoServicio = cmbTipoServicio.ListIndex
        oRptAHCEgresoMedico24.TextoDelFiltro = ml_TextoDelFiltro
        oRptAHCEgresoMedico24.CrearReporte_excel Me.hwnd
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   Tipo Servicio: " & cmbTipoServicio.Text & "     Responsable: " & cmbIdResponsable.Text & "     Orden: " & cmbOrden.Text & "     F.Alta Médica:(" & txtFechaInicio.Text & " hasta " & txtFechaFin.Text & ")"
    If cmbTipoServicio.Text = "" Then
        sMensaje = sMensaje + "Por favor elija el tipo de Servicio"
    End If
    
    If Me.txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de alta medica inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            sMensaje = "La fecha de alta medica inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFechaFin = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de alta medica final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            sMensaje = "La fecha de alta medica final no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub



Private Sub cmbIdResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsable
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable

End Sub


Private Sub Form_Load()
       
       mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
       mo_cmbIdResponsable.ListField = "ApAmNo"
       Set mo_cmbIdResponsable.RowSource = mo_AdminArchivoClinico.ArchiverosSeleccionarTodos()
       
       Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFechaFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       cmbTipoServicio.ListIndex = 0
       cmbOrden.ListIndex = 0
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
