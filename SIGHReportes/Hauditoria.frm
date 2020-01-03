VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Hauditoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auditoria Hospitalización"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   Icon            =   "Hauditoria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9195
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
      TabIndex        =   6
      Top             =   0
      Width           =   9195
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
         Left            =   7950
         Picture         =   "Hauditoria.frx":0CCA
         TabIndex        =   10
         Top             =   660
         Width           =   1125
      End
      Begin VB.ComboBox cmbUsuario 
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
         ItemData        =   "Hauditoria.frx":0FDC
         Left            =   1500
         List            =   "Hauditoria.frx":0FE6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   2145
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
         Left            =   6930
         TabIndex        =   1
         Top             =   210
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
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
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
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   8310
         TabIndex        =   12
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
         Left            =   6420
         TabIndex        =   9
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Usuario Sistema"
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
         Top             =   750
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   4
      Top             =   1950
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "Hauditoria.frx":1002
         DownPicture     =   "Hauditoria.frx":1462
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
         Left            =   3210
         Picture         =   "Hauditoria.frx":18D7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "Hauditoria.frx":1D4C
         DownPicture     =   "Hauditoria.frx":2210
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
         Left            =   4740
         Picture         =   "Hauditoria.frx":26FC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "Hauditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Auditoría en Hospitalización
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbUsuario As New SIGHEntidades.ListaDespleglable
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
Dim oRsTmp99 As New Recordset

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New RptHauditoria
        oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), "Auditoria en Hospitalización", ml_TextoDelFiltro, Val(mo_cmbUsuario.BoundText), CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS)), Me.hwnd
        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:    F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ")" & IIf(cmbUsuario.Text = "", "", "     Usuario: " & cmbUsuario.Text)
    
    If Me.txtFdesde = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento inicial"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = SIGHEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento final"
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento final no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrInicio = SIGHEntidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora de movimiento inicial"
    Else
        If Not SIGHEntidades.EsHora(txtHrInicio) Then
            sMensaje = "La hora de movimiento inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrFin = SIGHEntidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora de movimiento final"
    Else
        If Not SIGHEntidades.EsHora(txtHrFin) Then
            sMensaje = "La hora de movimiento final no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
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
    LimpiarVariablesDeMemoria
End Sub


Private Sub cmbUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
      mo_Teclado.RealizarNavegacion KeyCode, cmbUsuario
End Sub

Private Sub Form_Initialize()
    Set mo_cmbUsuario.MiComboBox = cmbUsuario
End Sub


Private Sub Form_Load()
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:01"
    txtHrFin.Text = "23:59"
    
    '
    mo_cmbUsuario.BoundColumn = "IdEmpleado"
    mo_cmbUsuario.ListField = "Usuario"
    Set mo_cmbUsuario.RowSource = mo_ReglasFarmacia.DevuelveListaDeUsuariosDelSistema
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
    Set mo_cmbUsuario = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
    If txtHrFin <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHrFin) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrFin = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub txtHrInicio_LostFocus()
    If txtHrInicio <> SIGHEntidades.HORA_VACIA_HM Then
        If Not SIGHEntidades.EsHora(txtHrInicio) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrInicio = SIGHEntidades.HORA_VACIA_HM
        End If
    End If
End Sub
