VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form rConsumoXservicio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consumo por Servicios"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rConsumoXservicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   60
      TabIndex        =   5
      Top             =   1800
      Width           =   6240
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rConsumoXservicio.frx":0CCA
         DownPicture     =   "rConsumoXservicio.frx":112A
         Height          =   700
         Left            =   1703
         Picture         =   "rConsumoXservicio.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rConsumoXservicio.frx":1A14
         DownPicture     =   "rConsumoXservicio.frx":1ED8
         Height          =   700
         Left            =   3233
         Picture         =   "rConsumoXservicio.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6270
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   1680
         TabIndex        =   11
         Top             =   960
         Width           =   4500
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         Height          =   315
         Left            =   90
         Picture         =   "rConsumoXservicio.frx":28B0
         TabIndex        =   8
         Top             =   1290
         Width           =   1785
      End
      Begin VB.ComboBox cmbConsiderar 
         Height          =   330
         ItemData        =   "rConsumoXservicio.frx":2BC2
         Left            =   1680
         List            =   "rConsumoXservicio.frx":2BCF
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   4500
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   600
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
         Left            =   4770
         TabIndex        =   9
         Top             =   600
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
         Caption         =   "Farmacia"
         Height          =   210
         Left            =   105
         TabIndex        =   12
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   210
         Left            =   4410
         TabIndex        =   10
         Top             =   630
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F. Movimiento"
         Height          =   210
         Left            =   105
         TabIndex        =   4
         Top             =   637
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Especialidad"
         Height          =   210
         Left            =   105
         TabIndex        =   3
         Top             =   285
         Width           =   1380
      End
   End
End
Attribute VB_Name = "rConsumoXservicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte Consumo por Servicio
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Dim mo_cmbAlmacen As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia

Private Sub btnAceptar_Click()
    If Me.txtFechaInicio.Text = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "Por favor ingrese la Fecha desde", vbInformation, Me.Caption
        txtFechaInicio.SetFocus
    ElseIf Me.txtFechaFin.Text = sighentidades.FECHA_VACIA_DMY Then
        MsgBox "Por favor ingrese la Fecha hasta", vbInformation, Me.Caption
        txtFechaFin.SetFocus
    Else
        If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
           Exit Sub
        End If
        Me.MousePointer = 11
        Dim oRptClaseCry As New rCrystal
        oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
        oRptClaseCry.FechaInicio = Format(txtFechaInicio.Text & " 00:00:01", "DD/MM/YYYY HH:MM:SS")
        oRptClaseCry.FechaFin = Format(Me.txtFechaFin.Text & " 23:59:59", "DD/MM/YYYY HH:MM:SS")
        oRptClaseCry.TipoServicioHosp = IIf(cmbConsiderar.ListIndex = 0, "3", IIf(cmbConsiderar.ListIndex = 1, "2", "1"))
        oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
        oRptClaseCry.TextoDelFiltro = "Farmacia: " & Trim(Me.cmbAlmacen.Text) & "  Fecha Movimiento: " & txtFechaInicio.Text & " al " & Me.txtFechaFin.Text & "     " & IIf(cmbConsiderar.ListIndex = 0, "(Hospitalización)", IIf(cmbConsiderar.ListIndex = 1, "(Emergencia)", "(Consultorios Externos)"))
        oRptClaseCry.TipoReporte = Me.Name
        oRptClaseCry.Show vbModal
        Set oRptClaseCry = Nothing
        Me.MousePointer = 1
    End If
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Sub InicializaFechaHora()
    Me.txtFechaInicio.Text = Date
    Me.txtFechaFin.Text = Date

End Sub

Private Sub Form_Load()
    cmbConsiderar.ListIndex = 0
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idtipoSuministro='01'")
    mo_cmbAlmacen.BoundText = "4"
End Sub
Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
End Sub

Private Sub txtFechaFin_LostFocus()
If Not sighentidades.esfecha(txtFechaFin.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtFechaInicio_LostFocus()
If Not sighentidades.esfecha(txtFechaInicio.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub
