VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CEatenciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atenciones vs Atendidos"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   1
      Top             =   2460
      Width           =   5760
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CEAtenciones.frx":0000
         DownPicture     =   "CEAtenciones.frx":0460
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
         Left            =   1470
         Picture         =   "CEAtenciones.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CEAtenciones.frx":0D4A
         DownPicture     =   "CEAtenciones.frx":120E
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
         Left            =   3000
         Picture         =   "CEAtenciones.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5730
      Begin VB.Frame fraFiltro 
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   2250
         TabIndex        =   9
         Top             =   150
         Visible         =   0   'False
         Width           =   3375
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
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   600
            Width           =   3225
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Médico"
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
            TabIndex        =   11
            Top             =   330
            Width           =   570
         End
      End
      Begin VB.CheckBox chkServicio 
         Caption         =   "Un solo Médico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1785
      End
      Begin MSMask.MaskEdBox txtFecha1 
         Height          =   315
         Left            =   1095
         TabIndex        =   4
         Top             =   1830
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
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   1830
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
      Begin VB.Label Label1 
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
         Left            =   2670
         TabIndex        =   8
         Top             =   1860
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Atención"
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
         Left            =   75
         TabIndex        =   5
         Top             =   1860
         Width           =   885
      End
   End
End
Attribute VB_Name = "CEatenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Atenciones vs Atendidos
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdResponsable As New SIGHEntidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptHistorias As New RptCEatenciones
        oRptHistorias.IdResponsable = Val(mo_cmbIdResponsable.BoundText)
        oRptHistorias.FechaInicio = txtFecha1.Text
        oRptHistorias.FechaFin = txtFecha2.Text
        oRptHistorias.TextoDelFiltro = "Fechas de Atención: " & txtFecha1.Text & " hasta " & txtFecha2.Text & IIf(chkServicio.Value = 1, "       Médico: " & cmbIdResponsable.Text, "")
        oRptHistorias.CrearReporte_excel Me.hwnd
        Me.MousePointer = 1
    End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If chkServicio.Value = 1 Then
        If mo_cmbIdResponsable.BoundText = "" Then
            sMensaje = sMensaje + "Por favor elija el Médico"
        End If
    End If
    
    If Me.txtFecha1 = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de atención inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFecha1, "DD/MM/AAAA") Then
            sMensaje = "La fecha de atención inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFecha2 = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de atención final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFecha2, "DD/MM/AAAA") Then
            sMensaje = "La fecha de atención final no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFecha1.Text) > CDate(Me.txtFecha2.Text) Then
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




Private Sub chkServicio_Click()
   If chkServicio.Value = 1 Then
      fraFiltro.Visible = True
   Else
      fraFiltro.Visible = False
   End If
End Sub



Private Sub Form_Initialize()
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable

End Sub


Private Sub Form_Load()
       Dim mo_AdminServHosp As New ReglasServiciosHosp
       Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica
       
       Me.txtFecha1.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFecha2.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       
       mo_cmbIdResponsable.BoundColumn = "IdMedico"
       mo_cmbIdResponsable.ListField = "Dmedico"
       Set mo_cmbIdResponsable.RowSource = oBuscaMedicos.MedicosSeleccionarTodosOrdenadoAlfabeticamente
       
       
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFecha1_LostFocus()
    If txtFecha1 <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecha1, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFecha1 = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub


Private Sub txtFecha2_LostFocus()
    If txtFecha2 <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFecha2, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFecha2 = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
