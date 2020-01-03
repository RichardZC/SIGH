VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form ReporteIngresosHosp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de ingresos "
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   17
      Top             =   6210
      Width           =   7635
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ReporteIngresosHosp.frx":0000
         DownPicture     =   "ReporteIngresosHosp.frx":0460
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
         Left            =   2423
         Picture         =   "ReporteIngresosHosp.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ReporteIngresosHosp.frx":0D4A
         DownPicture     =   "ReporteIngresosHosp.frx":120E
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
         Left            =   3953
         Picture         =   "ReporteIngresosHosp.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6165
      Left            =   30
      TabIndex        =   11
      Top             =   15
      Width           =   7635
      Begin VB.CheckBox chkDiario 
         Caption         =   "Reporte Diario Emergencia"
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
         Left            =   375
         TabIndex        =   36
         Top             =   4485
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtConsumo 
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
         Left            =   2535
         TabIndex        =   31
         Top             =   4110
         Width           =   795
      End
      Begin Threed.SSOption optIngresosH 
         Height          =   345
         Left            =   165
         TabIndex        =   23
         Top             =   195
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   609
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
         Caption         =   "Ingresos Hospitalarios"
      End
      Begin VB.ComboBox cmbConsiderar 
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
         ItemData        =   "ReporteIngresosHosp.frx":1BE6
         Left            =   2085
         List            =   "ReporteIngresosHosp.frx":1BF0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   675
         Width           =   5115
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipos de Número de Historia Clínica"
         Height          =   1065
         Left            =   405
         TabIndex        =   18
         Top             =   3015
         Width           =   6825
         Begin VB.ComboBox cmbTipoHistoria 
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
            ItemData        =   "ReporteIngresosHosp.frx":1C11
            Left            =   1635
            List            =   "ReporteIngresosHosp.frx":1C1E
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   230
            Width           =   5130
         End
         Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1635
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   630
            Visible         =   0   'False
            Width           =   5130
         End
         Begin VB.PictureBox XP_ProgressBar1 
            Height          =   300
            Left            =   135
            ScaleHeight     =   240
            ScaleWidth      =   5010
            TabIndex        =   19
            Top             =   2280
            Width           =   5070
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Considerar"
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
            TabIndex        =   20
            Top             =   285
            Width           =   840
         End
      End
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
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1095
         Width           =   5130
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
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1485
         Width           =   5130
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
         Left            =   2055
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1875
         Width           =   5145
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   2070
         TabIndex        =   4
         Top             =   2265
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
         Left            =   5055
         TabIndex        =   5
         Top             =   2265
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
      Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
         Height          =   330
         Left            =   2055
         TabIndex        =   6
         Top             =   2640
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optIngresosSIS 
         Height          =   345
         Left            =   75
         TabIndex        =   24
         Top             =   4860
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   609
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
         Caption         =   "Pacientes sin Alta Médica que pasaron más de 180 días de estancia (o por llegar a 180)"
         Value           =   -1
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   4350
         TabIndex        =   25
         Top             =   5310
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
      Begin MSMask.MaskEdBox txtNdias 
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   5295
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo cmdFuenteFinanc 
         Height          =   330
         Left            =   2025
         TabIndex        =   29
         Top             =   5685
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtHinicial 
         Height          =   315
         Left            =   3465
         TabIndex        =   34
         Top             =   2265
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
      Begin MSMask.MaskEdBox txtHfinal 
         Height          =   315
         Left            =   6450
         TabIndex        =   35
         Top             =   2265
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Consumo actual, mayor a"
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
         Left            =   420
         TabIndex        =   33
         Top             =   4140
         Width           =   2055
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(S/.)    (sin Alta Médica)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3420
         TabIndex        =   32
         Top             =   4155
         Width           =   2205
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fte.Financiam/IAFA"
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
         Left            =   435
         TabIndex        =   30
         Top             =   5730
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Días Estancia"
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
         Left            =   450
         TabIndex        =   28
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "hasta el"
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
         Left            =   3660
         TabIndex        =   26
         Top             =   5355
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fte.Financiam/IAFA"
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
         Left            =   435
         TabIndex        =   22
         Top             =   2685
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Especialidad"
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
         Left            =   420
         TabIndex        =   21
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   240
         Left            =   4500
         TabIndex        =   16
         Top             =   2325
         Width           =   480
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
         Left            =   420
         TabIndex        =   15
         Top             =   2295
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
         Left            =   420
         TabIndex        =   14
         Top             =   1140
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
         Left            =   420
         TabIndex        =   13
         Top             =   1515
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
         Left            =   420
         TabIndex        =   12
         Top             =   1920
         Width           =   1275
      End
   End
End
Attribute VB_Name = "ReporteIngresosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Ingresos Hospitalarios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_cmbIdDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbIdServicio As New sighentidades.ListaDespleglable
Dim mo_cmbIdEspecialidad As New sighentidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_IdTipoReporte As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New sighentidades.ListaDespleglable
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim lcFiltro As String
Dim oRsFuentesFinanciamiento As New Recordset
Dim mo_MostrarReporte As Boolean
Property Let mostrarReporte(lIdValue As Boolean)
    mo_MostrarReporte = lIdValue
    chkDiario.Visible = False
    If mo_MostrarReporte = True Then
       btnAceptar_Click
    Else
       If cmbConsiderar.ListIndex = 1 Then chkDiario.Visible = True
    End If
End Property

Property Let IdTipoReporte(lIdValue As Long)
    ml_IdTipoReporte = lIdValue
    
End Property

Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

    Dim oRptIngresosHosp As New clReporteIngrHosp
    Me.MousePointer = 11
    If optIngresosH.Value = True Then
        If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de inicio", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
                MsgBox "La fecha de inicio, no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFechaFin = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha final", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
                MsgBox "La fecha final, no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Sub
        End If
        
        lcFiltro = "Filtros:  F.Ingreso: (" & txtFechaInicio.Text & " " & txtHinicial.Text & " - " & _
                   txtFechaFin.Text & " " & txtHfinal.Text & ") " & _
                 "     (" & cmbConsiderar.Text & ")     " & _
                 IIf(cmbIdDepartamento.Text = "", "", "     Departamento: " & cmbIdDepartamento.Text) & _
                 IIf(cmbIdEspecialidad.Text = "", "", "     Especialidad: " & cmbIdEspecialidad.Text) & _
                 IIf(cmbIdServicio.Text = "", "", "     Servicio: " & cmbIdServicio.Text) & _
                 IIf(Val(txtConsumo.Text) = 0, "", "     Cosumos mayor a: S/ " & Trim(txtConsumo.Text) & " sin ALTA MEDICA")
        
        Select Case ml_IdTipoReporte
        Case sghReporteIngresosHospitalario
            'Dim oRptIngresosHosp As New RptIngresosHosp
            
            
            oRptIngresosHosp.IdDepartamento = Val(mo_cmbIdDepartamento.BoundText)
            oRptIngresosHosp.IdEspecialidad = Val(mo_cmbIdEspecialidad.BoundText)
            oRptIngresosHosp.idServicio = Val(mo_cmbIdServicio.BoundText)
            oRptIngresosHosp.FechaFin = Me.txtFechaFin.Text
            oRptIngresosHosp.FechaInicio = Me.txtFechaInicio.Text
            'Set oRptIngresosHosp.progressRpt = Me.progressRpt
            oRptIngresosHosp.IdTipoNroHistoria = IIf(cmbTipoHistoria.ListIndex = 2, mo_cmbIdTipoGenHistoriaClinica.BoundText, IIf(cmbTipoHistoria.ListIndex = 0, 100, 200))
            oRptIngresosHosp.IdTipoEspecialidad = IIf(cmbConsiderar.ListIndex = 0, 3, 2)
            oRptIngresosHosp.TextoDelFiltro = lcFiltro
            oRptIngresosHosp.TextoDelFiltro = lcFiltro + IIf(Val(cmbFuenteFinanciamiento.BoundText) > 0, "  (IAFA: " & Trim(cmbFuenteFinanciamiento.Text) & ")", "")
            oRptIngresosHosp.IdPlan = Val(cmbFuenteFinanciamiento.BoundText)
            If Me.chkDiario.Value = 1 Then
               oRptIngresosHosp.CrearReporteIngresosEmergencia Me.hwnd, Val(Me.txtConsumo.Text), txtHinicial.Text, txtHfinal.Text
            Else
               oRptIngresosHosp.CrearReporteIngresosHospitalarios Me.hwnd, Val(Me.txtConsumo.Text), txtHinicial.Text, txtHfinal.Text
            End If
        Case 2
        End Select
    Else
        oRptIngresosHosp.FechaFin = Me.txtFhasta.Text
        oRptIngresosHosp.FechaInicio = "01/01/2000"
        oRptIngresosHosp.TextoDelFiltro = "F.Ingreso hasta: " & txtFhasta.Text & _
                                          "  (Fuente Financiamiento: " & Trim(cmdFuenteFinanc.Text) & ")"
        oRptIngresosHosp.IdPlan = Val(cmdFuenteFinanc.BoundText)
        oRptIngresosHosp.CrearReportePacientesSISconMas180diasEstancia Val(txtNdias.Text), Me.hwnd
    End If
    Me.MousePointer = 1
    Set oRptIngresosHosp = Nothing
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
        MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdEspecialidad_Click()
    
    mo_cmbIdServicio.BoundColumn = "IdServicio"
    mo_cmbIdServicio.ListField = "DescripcionLarga"
    If cmbConsiderar.ListIndex = 0 Then
       Set mo_cmbIdServicio.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento.BoundText), Val(mo_cmbIdEspecialidad.BoundText))
    Else
       Set mo_cmbIdServicio.RowSource = mo_reglasComunes.ServiciosSeleccionarEmergenciaPorEspecialidad(Val(mo_cmbIdEspecialidad.BoundText))
    End If

End Sub


Private Sub cmbTipoHistoria_Change()
    cmbTipoHistoria_Click
End Sub

Private Sub cmbTipoHistoria_Click()
   If cmbTipoHistoria.ListIndex = 2 Then
      cmbIdTipoGenHistoriaClinica.Visible = True
   Else
      cmbIdTipoGenHistoriaClinica.Visible = False
   End If

End Sub

Private Sub Form_Initialize()

    Set mo_cmbIdDepartamento.MiComboBox = cmbIdDepartamento
    Set mo_cmbIdEspecialidad.MiComboBox = cmbIdEspecialidad
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica

    Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    Me.txtFechaFin = sighentidades.UltimaFechaDDMMYYDelMesActual()
    
End Sub

Private Sub Form_Load()
       

        Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
        Me.txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
        txtHinicial.Text = "00:00"
        txtHfinal.Text = "23:59"
        
        cmbTipoHistoria.ListIndex = 0
    
        mo_cmbIdDepartamento.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamento.ListField = "DescripcionLarga"
        Set mo_cmbIdDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
        
        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        mo_cmbIdTipoGenHistoriaClinica.BoundText = 2
        
        Set oRsFuentesFinanciamiento = mo_reglasComunes.FuentesFinanciamientoSegunFiltro("")
       Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
       cmbFuenteFinanciamiento.ListField = "Descripcion"
       cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
       
       cmbConsiderar.ListIndex = 0
       
       txtNdias.Text = Val(lcBuscaParametro.SeleccionaFilaParametro(360))
       txtFhasta.Text = Format(Date + 2, sighentidades.DevuelveFechaSoloFormato_DMY)
       Set cmdFuenteFinanc.RowSource = oRsFuentesFinanciamiento
       cmdFuenteFinanc.ListField = "Descripcion"
       cmdFuenteFinanc.BoundColumn = "idFuenteFinanciamiento"
       cmdFuenteFinanc.BoundText = Trim(Str(sghFuenteFinanciamiento.sghFFSIS))
       
       
       
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

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

