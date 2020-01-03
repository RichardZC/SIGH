VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form rptCajaDevoluciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caja Devoluciones"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "rptCajaDevoluciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   9240
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
      Height          =   1395
      Left            =   0
      TabIndex        =   5
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
         Left            =   7080
         Picture         =   "rptCajaDevoluciones.frx":0CCA
         TabIndex        =   8
         Top             =   600
         Width           =   1125
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rptCajaDevoluciones.frx":0FDC
         DownPicture     =   "rptCajaDevoluciones.frx":143C
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
         Picture         =   "rptCajaDevoluciones.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rptCajaDevoluciones.frx":1D26
         DownPicture     =   "rptCajaDevoluciones.frx":21EA
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
         Picture         =   "rptCajaDevoluciones.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rptCajaDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Auditoría en Archivo Clínico
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbUsuario As New sighEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New sighEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New sighEntidades.Formulario
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New clRptCajaDevoluciones
        oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), "Caja Devoluciones", ml_TextoDelFiltro, CDate(Format(txtFdesde.Text, sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)), CDate(Format(txtFhasta.Text, sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)), Me.hwnd
        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:    F.Movimiento: (" & txtFdesde.Text & "   al " & txtFhasta.Text & " ) )"
    
    If Me.txtFdesde = sighEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento inicial"
    Else
        If Not sighEntidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = sighEntidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento final"
    Else
        If Not sighEntidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento final no tiene el formato correcto"
        End If
    End If
    
    If CDate(Me.txtFdesde.Text) > CDate(Me.txtFhasta.Text) Then
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


Private Sub Form_Load()
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    
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
    If txtFdesde <> sighEntidades.FECHA_VACIA_DMY Then
        If Not sighEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde = sighEntidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> sighEntidades.FECHA_VACIA_DMY Then
        If Not sighEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFhasta = sighEntidades.FECHA_VACIA_DMY
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
    Set mo_reglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub


