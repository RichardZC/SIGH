VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptHISDxOmitidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de diagnósticos omitidos"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   5130
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "RptHISDxOmitidos.frx":0000
         DownPicture     =   "RptHISDxOmitidos.frx":0460
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
         Left            =   1200
         Picture         =   "RptHISDxOmitidos.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "RptHISDxOmitidos.frx":0D4A
         DownPicture     =   "RptHISDxOmitidos.frx":120E
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
         Left            =   2730
         Picture         =   "RptHISDxOmitidos.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      Begin VB.ComboBox cmbMes 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   135
         ScaleHeight     =   240
         ScaleWidth      =   5010
         TabIndex        =   5
         Top             =   2280
         Width           =   5070
      End
      Begin MSMask.MaskEdBox mskfechaAnio 
         Height          =   330
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Año"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Mes"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "RptHISDxOmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Diagnósticos Omitidos (HIS)
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado

Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptHisDxOmitidos As New clRptHisDxOmitidos
        oRptHisDxOmitidos.Mes = Val(mo_cmbMes.BoundText)
        oRptHisDxOmitidos.Anio = Me.mskfechaAnio.Text
        oRptHisDxOmitidos.Texto = "Reporte de " & Me.cmbMes.Text & " " & Me.mskfechaAnio.Text
        oRptHisDxOmitidos.CrearReporte_excel Me.hwnd
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If Trim(Me.cmbMes.Text) = "" Then
        sMensaje = sMensaje + "Debe seleccionar una mes válido." + vbCrLf
    End If
    If Trim(Me.mskfechaAnio.Text) = "____" Or IsNumeric(Me.mskfechaAnio.Text) = False Then
        sMensaje = sMensaje + "Debe ingresar una año válido." + vbCrLf
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

Private Sub cmbmes_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbMes
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbMes.MiComboBox = cmbMes
End Sub

Private Sub Form_Load()
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
    mo_cmbMes.BoundText = Month(Now)
    mskfechaAnio.Text = Year(Now)
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub mskfechaAnio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, mskfechaAnio
    AdministrarKeyPreview KeyCode
End Sub


