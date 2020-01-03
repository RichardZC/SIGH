VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form HAlojados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alojados"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "HAlojados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5595
      Begin VB.ComboBox cmbServicioIngreso 
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
         Left            =   1500
         TabIndex        =   10
         Top             =   660
         Width           =   3960
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Picture         =   "HAlojados.frx":0CCA
         TabIndex        =   9
         Top             =   1050
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   240
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
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   4110
         TabIndex        =   1
         Top             =   240
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
      Begin VB.Label Label4 
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
         Left            =   3600
         TabIndex        =   8
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Ingreso"
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
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   3
      Top             =   1950
      Width           =   5610
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HAlojados.frx":0FDC
         DownPicture     =   "HAlojados.frx":143C
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
         Left            =   1373
         Picture         =   "HAlojados.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HAlojados.frx":1D26
         DownPicture     =   "HAlojados.frx":21EA
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
         Left            =   2903
         Picture         =   "HAlojados.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "HAlojados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Pacientes alojados
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbServicioIngreso As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long
Dim ml_IdEspecialidad As Long
Const lnIdEspecialidad As Long = 23    'Neonatologia

Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New RptHAlojados
        oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), "Alojados", ml_TextoDelFiltro, Val(mo_cmbServicioIngreso.BoundText), CDate(txtFdesde.Text), CDate(txtFhasta.Text), Me.hwnd
        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    
    If Me.txtFdesde = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de ingreso inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha de ingreso inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de ingreso final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha de ingreso final no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFdesde.Text) > CDate(Me.txtFhasta.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
    End If
    
    ml_TextoDelFiltro = "FILTROS:    F.Ingreso: (" & txtFdesde.Text & "  al " & txtFhasta.Text & ")" & IIf(cmbServicioIngreso.Text = "", "", "     Servicio: " & cmbServicioIngreso.Text)
    If cmbServicioIngreso.Text = "" Then
       sMensaje = "Por favor elija el SERVICIO" + Chr(13)
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
    LimpiarVariablesDeMemoria
End Sub




Private Sub Form_Initialize()
    Set mo_cmbServicioIngreso.MiComboBox = cmbServicioIngreso
End Sub


Private Sub Form_Load()
    txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual()
    txtFhasta.Text = Date
    
    Dim oBuscaServicios As New SIGHNegocios.ReglasAdmision
    Dim lcEspecialidadesDelUsuario As String
    lcEspecialidadesDelUsuario = oBuscaServicios.DevuelveEspecialidadesServicioSegunUsuarioSistema(sghEspecialidadesHosp, ml_idUsuario)
    lcEspecialidadesDelUsuario = " and dbo.Servicios.IdEspecialidad= " & lnIdEspecialidad
    
    mo_cmbServicioIngreso.BoundColumn = "IdServicio"
    mo_cmbServicioIngreso.ListField = "DservicioHosp"
    Set mo_cmbServicioIngreso.RowSource = oBuscaServicios.DevuelveServiciosDelHospital("(3)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
    Set oBuscaServicios = Nothing
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub






Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFhasta = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbServicioIngreso = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub


