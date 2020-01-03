VERSION 5.00
Begin VB.Form HerrAperturaAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apertura Anual"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HerrAperturaAnual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5325
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
      Height          =   1065
      Left            =   30
      TabIndex        =   2
      Top             =   3990
      Width           =   5235
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrAperturaAnual.frx":0CCA
         DownPicture     =   "HerrAperturaAnual.frx":118E
         Height          =   700
         Left            =   2745
         Picture         =   "HerrAperturaAnual.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HerrAperturaAnual.frx":1B66
         DownPicture     =   "HerrAperturaAnual.frx":1FC6
         Height          =   700
         Left            =   1245
         Picture         =   "HerrAperturaAnual.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones:"
      Height          =   3885
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      Begin VB.ComboBox cmbAnios 
         Height          =   330
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3450
         Width           =   1035
      End
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         ForeColor       =   &H80000004&
         Height          =   3000
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4965
      End
      Begin VB.Label Label1 
         Caption         =   "Año a Aperturar"
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   3480
         Width           =   1365
      End
   End
End
Attribute VB_Name = "HerrAperturaAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Herramienta para apertura ANUAL
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim lcDias As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lcNombrePc  As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_ReglasComunes As New ReglasComunes
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property


Private Sub btnAceptar_Click()
    If cmbAnios.Text = "" Then
       MsgBox "Elija el año a aperturar", vbInformation, "Apertura"
       Exit Sub
    End If
    If MsgBox("Esta seguro que desea APERTURAR EL AÑO: " & cmbAnios.Text, vbQuestion + vbYesNo, "Farmacia") = vbYes Then
        Me.MousePointer = 1
        mo_ReglasFarmacia.ActualizaCorrelativosPorAnio (cmbAnios.Text)
        '
        Me.MousePointer = 11
        Me.Visible = False
        LimpiarVariablesDeMemoria
    End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Private Sub Form_Load()
  mo_ReglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrAperturaAnual"
  cmbAnios.AddItem Trim(Str(Val(lcBuscaParametro.SeleccionaFilaParametro(228)) + 1))
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set lcBuscaParametro = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub
