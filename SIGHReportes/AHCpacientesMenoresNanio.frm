VERSION 5.00
Begin VB.Form AHCpacientesMenoresNanio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historias Clínicas de Pacientes Menores a N años"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   5400
      Begin VB.PictureBox progressRpt 
         Height          =   300
         Left            =   4050
         ScaleHeight     =   240
         ScaleWidth      =   1020
         TabIndex        =   6
         Top             =   1050
         Visible         =   0   'False
         Width           =   1080
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
         Left            =   210
         Picture         =   "AHCpacientesMenoresNanio.frx":0000
         TabIndex        =   5
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox txtAnios 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2220
         TabIndex        =   4
         Text            =   "3"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblIdTipoHistoria 
         Caption         =   "Pacientes menores a "
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
         Left            =   210
         TabIndex        =   8
         Top             =   285
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "años"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2730
         TabIndex        =   7
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   0
      Top             =   1770
      Width           =   5400
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHCpacientesMenoresNanio.frx":0312
         DownPicture     =   "AHCpacientesMenoresNanio.frx":0772
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
         Left            =   1320
         Picture         =   "AHCpacientesMenoresNanio.frx":0BE7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHCpacientesMenoresNanio.frx":105C
         DownPicture     =   "AHCpacientesMenoresNanio.frx":1520
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
         Left            =   2850
         Picture         =   "AHCpacientesMenoresNanio.frx":1A0C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "AHCpacientesMenoresNanio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Pacientes menores a NN años
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New SIGHEntidades.Teclado



Private Sub btnAceptar_Click()
    If Val(Me.txtAnios.Text) <= 0 Then
        MsgBox "Por favor ingresar la edad", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = 11
    Dim oRptAHCpacienteHastaNanio As New RptAHCpacienteHastaNanio
    oRptAHCpacienteHastaNanio.EdadMaxima = Me.txtAnios.Text
    oRptAHCpacienteHastaNanio.TextoDelFiltro = "Filtros:  Pacientes menores a : " & Me.txtAnios.Text
    oRptAHCpacienteHastaNanio.CrearReporte IIf(chkExcel.Value = 1, True, False), True, "", "", Me.hwnd
    Me.MousePointer = 1
    
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub







