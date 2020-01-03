VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HIndiHospConsolidado 
   Caption         =   "Reporte de ingresos de hospitalizaciòn"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   12
      Top             =   2265
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HIndiHospConsolidado.frx":0000
         DownPicture     =   "HIndiHospConsolidado.frx":0460
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
         Picture         =   "HIndiHospConsolidado.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HIndiHospConsolidado.frx":0D4A
         DownPicture     =   "HIndiHospConsolidado.frx":120E
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
         Picture         =   "HIndiHospConsolidado.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5370
      Begin VB.ComboBox cmbIdDepartamento 
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
         TabIndex        =   3
         Top             =   230
         Width           =   3555
      End
      Begin VB.ComboBox cmbIdEspecialidad 
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
         TabIndex        =   2
         Top             =   620
         Width           =   3555
      End
      Begin VB.ComboBox cmbIdServicio 
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
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1010
         Width           =   3570
      End
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   135
         ScaleHeight     =   240
         ScaleWidth      =   5010
         TabIndex        =   4
         Top             =   2280
         Width           =   5070
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1395
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1755
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
      Begin VB.Label Label3 
         Caption         =   "Fecha Ingreso Fin"
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
         Left            =   165
         TabIndex        =   11
         Top             =   1815
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Ingreso Ini."
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
         Left            =   150
         TabIndex        =   10
         Top             =   1440
         Width           =   1560
      End
      Begin VB.Label Departamento 
         Caption         =   "Dpto ingreso"
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
         Left            =   150
         TabIndex        =   9
         Top             =   285
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   "Esp. ingreso"
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
         Left            =   150
         TabIndex        =   8
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Serv. ingreso"
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
         TabIndex        =   7
         Top             =   1065
         Width           =   1275
      End
   End
End
Attribute VB_Name = "HIndiHospConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mo_cmbIdDepartamento As New SIGHComun.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHComun.ListaDespleglable
Dim mo_cmbIdEspecialidad As New SIGHComun.ListaDespleglable

Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_Teclado As New SIGHComun.Teclado
Dim ml_IdTipoReporte As Long

Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
End Property

Private Sub btnAceptar_Click()
    If mo_cmbIdDepartamento.BoundText = "" Then
        MsgBox "Por favor ingrese el departamento", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If txtFechaInicio.Text = SIGHComun.FECHA_VACIA_DMY Then
        MsgBox "Por favor ingrese la fecha de inicio", vbInformation, Me.Caption
        Exit Sub
    End If

    If txtFechaFin.Text = SIGHComun.FECHA_VACIA_DMY Then
        MsgBox "Por favor ingrese la fecha final", vbInformation, Me.Caption
        Exit Sub
    End If

    Select Case ml_IdTipoReporte
    Case sghReporteIngresosHospitalario
'        Dim oRptIngresosHosp As New RptIngresosHosp
'        oRptIngresosHosp.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
'        oRptIngresosHosp.IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
'        oRptIngresosHosp.IdServicio = Val(mo_cmbIdServicio.BoundText)
'        oRptIngresosHosp.FechaFin = Me.txtFechaFin.Text
'        oRptIngresosHosp.FechaInicio = Me.txtFechaInicio.Text
'        Set oRptIngresosHosp.progressRpt = Me.progressRpt
'        oRptIngresosHosp.CrearReporteIngresosHospitalarios
    Case 2
    End Select

End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub cmbIdDepartamento_Click()
Dim sMensaje As String

       mo_cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento.BoundText))
       
       mo_cmbIdEspecialidad.BoundText = ""
       
       If mo_AdminServiciosHosp.MensajeError <> "" Then
        MsgBox mo_AdminServiciosHosp.MensajeError, vbCritical, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdEspecialidad_Click()
    
    mo_cmbIdServicio.BoundColumn = "IdServicio"
    mo_cmbIdServicio.ListField = "DescripcionLarga"
    Set mo_cmbIdServicio.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento.BoundText), Val(mo_cmbIdEspecialidad.BoundText))

End Sub

Private Sub Form_Initialize()

    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    
    Me.txtFechaInicio.Text = SIGHComun.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin = SIGHComun.UltimaFechaDDMMYYDelMesActual()
    
End Sub

Private Sub Form_Load()
       

        Me.txtFechaInicio.Text = SIGHComun.PrimerFechaDDMMYYDelMesActual()
        Me.txtFechaFin.Text = Format(Date, "dd/mm/yyyy")
    
       mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()

End Sub

Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaFin
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFechaInicio
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


