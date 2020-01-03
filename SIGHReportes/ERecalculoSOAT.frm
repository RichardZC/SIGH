VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ERecalculoSOAT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pacientes SOAT que pasaron a PARTICULAR"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   1
      Top             =   2100
      Width           =   5610
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ERecalculoSOAT.frx":0000
         DownPicture     =   "ERecalculoSOAT.frx":0460
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
         Picture         =   "ERecalculoSOAT.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ERecalculoSOAT.frx":0D4A
         DownPicture     =   "ERecalculoSOAT.frx":120E
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
         Left            =   2963
         Picture         =   "ERecalculoSOAT.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2010
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5640
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Height          =   315
         Left            =   90
         Picture         =   "ERecalculoSOAT.frx":1BE6
         TabIndex        =   8
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin MSMask.MaskEdBox txtFecha1 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   690
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
         Left            =   4140
         TabIndex        =   5
         Top             =   660
         Width           =   1365
         _ExtentX        =   2408
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
         Left            =   1500
         TabIndex        =   9
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Plan"
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
         TabIndex        =   10
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F.Egreso Médico"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1335
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
         Left            =   3900
         TabIndex        =   6
         Top             =   690
         Width           =   120
      End
   End
End
Attribute VB_Name = "ERecalculoSOAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Recalculo del SOAT
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_idUsuarioConPermisoEnSISoEXOoSOAT As Long
Dim ml_idUsuario As Long
Dim oRsFuentesFinanciamiento As New Recordset
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property



Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New RptEconRecalculoSOAT
        oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), Me.Caption, "F.Alta Médica: " & Me.txtFecha1.Text & " al " & Me.txtFecha2.Text, Val(Me.cmbFuenteFinanciamiento.BoundText), CDate(Me.txtFecha1.Text), CDate(Me.txtFecha2.Text), Me.hwnd
        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    
    If Me.txtFecha1 = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de egreso médico inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFecha1, "DD/MM/AAAA") Then
            sMensaje = "La fecha de egreso médico inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFecha2 = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de egreso médico final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFecha2, "DD/MM/AAAA") Then
            sMensaje = "La fecha de egreso médico final no tiene el formato correcto"
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


Private Sub Form_Load()
       Me.txtFecha1.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
       Me.txtFecha2.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       '
       Set oRsFuentesFinanciamiento = mo_ReglasFacturacion.FuentesFinanciamientoSeleccionarTodos
       Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
       cmbFuenteFinanciamiento.ListField = "Descripcion"
       cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
       cmbFuenteFinanciamiento.BoundText = "2"
       
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
