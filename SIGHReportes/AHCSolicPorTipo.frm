VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form AHCSolicPorTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historias Clínicas Solicitadas según Tipo Historia"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   45
      TabIndex        =   4
      Top             =   2175
      Width           =   5370
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCSolicPorTipo.frx":0000
         DownPicture     =   "AHCSolicPorTipo.frx":0460
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
         Picture         =   "AHCSolicPorTipo.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCSolicPorTipo.frx":0D4A
         DownPicture     =   "AHCSolicPorTipo.frx":120E
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
         Picture         =   "AHCSolicPorTipo.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5370
      Begin VB.CheckBox chkMuestraHistorial 
         Alignment       =   1  'Right Justify
         Caption         =   "Muestra Historial"
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
         Left            =   135
         Picture         =   "AHCSolicPorTipo.frx":1BE6
         TabIndex        =   13
         Top             =   1410
         Width           =   1725
      End
      Begin VB.ComboBox cmbIdTipoHistoria 
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
         TabIndex        =   7
         Top             =   210
         Width           =   3540
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
         Left            =   1665
         TabIndex        =   9
         Top             =   1035
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
         Left            =   3810
         TabIndex        =   10
         Top             =   1005
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
         Left            =   3315
         TabIndex        =   12
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Requerimiento"
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
         Left            =   150
         TabIndex        =   11
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Label lblIdTipoHistoria 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Historia"
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
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   660
         Width           =   1005
      End
   End
End
Attribute VB_Name = "AHCSolicPorTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Historias Solicitadas por Tipo
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdTipoHistoria As New sighentidades.ListaDespleglable
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptHistorias As New RptAHSolicPorTipo
        oRptHistorias.IdResponsable = Val(mo_cmbIdResponsable.BoundText)
        oRptHistorias.IdTipoHistoria = Val(mo_cmbIdTipoHistoria.BoundText)
        oRptHistorias.FechaInicio = txtFechaInicio.Text
        oRptHistorias.FechaFin = txtFechaFin.Text
        oRptHistorias.TextoDelFiltro = "Tipo de Historias: " & cmbIdTipoHistoria.Text & "     Responsable : " & cmbIdResponsable.Text & "     F.Requerimiento:(" & txtFechaInicio.Text & " al " & txtFechaFin.Text & ")"
        oRptHistorias.CrearReporte_excel Me.hwnd, IIf(chkMuestraHistorial.Value = 1, True, False)
        Me.MousePointer = 1
    End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If mo_cmbIdTipoHistoria.BoundText = "" Then
        sMensaje = sMensaje + "Por favor elija el Tipo de Historia"
    End If
    
    If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de requerimiento inicial"
    Else
        If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            sMensaje = "La fecha de requerimiento inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFechaFin = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de requerimiento final"
    Else
        If Not sighentidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            sMensaje = "La fecha de requerimiento final no tiene el formato correcto"
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
    Set mo_cmbIdTipoHistoria.MiComboBox = cmbIdTipoHistoria
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable

End Sub


Private Sub Form_Load()
       mo_cmbIdTipoHistoria.BoundColumn = "IdTipoHistoria"
       mo_cmbIdTipoHistoria.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoHistoria.RowSource = mo_AdminArchivoClinico.TiposHistoriaClinicaSeleccionarTodos()
       mo_cmbIdTipoHistoria.BoundText = 1
       sMensaje = sMensaje + mo_AdminArchivoClinico.MensajeError
       
       mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
       mo_cmbIdResponsable.ListField = "ApAmNo"
       Set mo_cmbIdResponsable.RowSource = mo_AdminArchivoClinico.ArchiverosSeleccionarTodos()
       
       Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
       mo_cmbIdTipoHistoria.BoundText = "4"
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
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
