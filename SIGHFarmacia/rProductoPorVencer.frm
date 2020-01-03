VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form rProductoPorVencer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Productos por Vencer"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rProductoPorVencer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   9255
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
      Width           =   9195
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         Height          =   315
         Left            =   60
         Picture         =   "rProductoPorVencer.frx":0CCA
         TabIndex        =   8
         Top             =   540
         Width           =   1125
      End
      Begin VB.TextBox txtDias 
         Height          =   315
         Left            =   3390
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   525
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   5370
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
      Begin VB.Label Label1 
         Caption         =   "días a partir del "
         Height          =   210
         Left            =   4020
         TabIndex        =   7
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Medicamentos que están por vencer a"
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   3300
      End
   End
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
      Height          =   1110
      Left            =   0
      TabIndex        =   3
      Top             =   1950
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rProductoPorVencer.frx":0FDC
         DownPicture     =   "rProductoPorVencer.frx":143C
         Height          =   700
         Left            =   3210
         Picture         =   "rProductoPorVencer.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rProductoPorVencer.frx":1D26
         DownPicture     =   "rProductoPorVencer.frx":21EA
         Height          =   700
         Left            =   4740
         Picture         =   "rProductoPorVencer.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "rProductoPorVencer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Producto por Vencer
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_mostrarReporte As Boolean

Property Let mostrarReporte(lValue As Boolean)
    ml_mostrarReporte = lValue
End Property

Private Sub btnAceptar_Click()
        Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.Dias = Val(txtDias.Text)
            oRptClaseCry.FechaInicio = txtFdesde.Text
            oRptClaseCry.TextoDelFiltro = "Productos por Vencer hasta el " & CDate(txtFdesde.Text) + Val(txtDias.Text)
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
        Me.MousePointer = 1
End Sub

Private Sub Form_Load()
    txtFdesde.Text = Date
    txtDias.Text = lcBuscaParametro.SeleccionaFilaParametro(220)
    If ml_mostrarReporte = True Then
       btnAceptar_Click
    End If
End Sub
Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub


Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtDias_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDias

End Sub



Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub

Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.esfecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde = Date
        End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
'           ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
End Sub

