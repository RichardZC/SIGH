VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form CErepMaterno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes para el Módulo Materno "
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "CErepMaterno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      Caption         =   "Reporte de seguimiento de gestantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9075
      Begin VB.TextBox txtMaxControles 
         Alignment       =   2  'Center
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "6"
         Top             =   360
         Width           =   585
      End
      Begin Threed.SSOption opcReporte2 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
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
         Caption         =   "Programadas para parto, FPP desde"
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   360
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
      Begin Threed.SSOption opcReporte1 
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
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
         Caption         =   "Con menos de"
         Value           =   -1
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   960
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
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
         Left            =   5400
         TabIndex        =   10
         Top             =   960
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
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
         Left            =   4920
         TabIndex        =   11
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "controles hasta la fecha"
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
         Left            =   2520
         TabIndex        =   7
         Top             =   405
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   9075
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CErepMaterno.frx":0CCA
         DownPicture     =   "CErepMaterno.frx":112A
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
         Picture         =   "CErepMaterno.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CErepMaterno.frx":1A14
         DownPicture     =   "CErepMaterno.frx":1ED8
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
         Picture         =   "CErepMaterno.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CErepMaterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte para módulo Materno
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros

Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Dim oRpt As New clCeMaterno
        If Me.opcReporte1.Value = True Then
            Me.MousePointer = 11
            oRpt.CrearReporteSeguimientoGestantes Val(Me.txtMaxControles.Text), Me.txtFecha.Text, Me.hwnd
            Me.MousePointer = 1
        End If
        If Me.opcReporte2.Value = True Then
            Me.MousePointer = 11
            oRpt.CrearReporteGestantesProgramasParto Me.txtFechaInicio.Text, Me.txtFechaFin.Text, Me.hwnd
            Me.MousePointer = 1
        End If
        Set oRpt = Nothing
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    ValidaDatosObligatorios = False
    If Me.opcReporte1.Value = True Then
        If Me.txtMaxControles.Text = "" Then
           MsgBox "Ingrese el número de controles", vbInformation, Me.Caption
           Exit Function
        End If
        If Me.txtFecha.Text = SIGHEntidades.FECHA_VACIA_DMY Then
           MsgBox "Ingrese la fecha", vbInformation, Me.Caption
           Exit Function
        End If
        If Not IsDate(Me.txtFecha.Text) Then
           MsgBox "La fecha no es válida", vbInformation, Me.Caption
           Exit Function
        End If
    End If
    If Me.opcReporte2.Value = True Then
        If Me.txtFechaInicio.Text = SIGHEntidades.FECHA_VACIA_DMY Then
           MsgBox "Ingrese la fecha desde", vbInformation, Me.Caption
           Exit Function
        End If
        If Not IsDate(Me.txtFechaInicio.Text) Then
           MsgBox "La fecha desde no es válida", vbInformation, Me.Caption
           Exit Function
        End If
        If Me.txtFechaFin.Text = SIGHEntidades.FECHA_VACIA_DMY Then
           MsgBox "Ingrese la fecha hasta", vbInformation, Me.Caption
           Exit Function
        End If
        If Not IsDate(Me.txtFechaFin.Text) Then
           MsgBox "La fecha hasta no es válida", vbInformation, Me.Caption
           Exit Function
        End If
        If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
           MsgBox "La fecha desde no puede ser mayor a la fecha hasta", vbInformation, Me.Caption
           Exit Function
        End If
    End If
    ValidaDatosObligatorios = True
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub Form_Load()
    txtFecha.Text = lcBuscaParametro.RetornaFechaServidorSQL
    txtFechaInicio.Text = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + 7
    txtFechaFin.Text = CDate(lcBuscaParametro.RetornaFechaServidorSQL) + 14
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtMaxControles_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
            KeyAscii = 0
        End If
    End If
End Sub

