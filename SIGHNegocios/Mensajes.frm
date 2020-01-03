VERSION 5.00
Begin VB.Form Mensajes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   Icon            =   "Mensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "Mensajes.frx":0CCA
      DownPicture     =   "Mensajes.frx":118E
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
      Left            =   5640
      Picture         =   "Mensajes.frx":167A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      DisabledPicture =   "Mensajes.frx":1B66
      DownPicture     =   "Mensajes.frx":1FC6
      Height          =   700
      Left            =   2040
      Picture         =   "Mensajes.frx":243B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "Mensajes.frx":28B0
      DownPicture     =   "Mensajes.frx":2D74
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
      Left            =   3840
      Picture         =   "Mensajes.frx":3260
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1365
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   1695
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   9375
   End
End
Attribute VB_Name = "Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mensajes de advertencia en para usar en Formulario
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_Mensaje As String
Dim ml_LongitudLetra As Integer
Dim ml_Titulo As String
Dim ml_EsNegrita As Boolean
Dim ml_ColorLetra As sghColores
Dim ml_UsaBotonesAceptarCancelar As Boolean
Dim mi_BotonPresionado As sghBotonDetallePresionado

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let UsaBotonesAceptarCancelar(lValue As Boolean)
    ml_UsaBotonesAceptarCancelar = lValue
End Property


Property Let ColorLetra(lValue As sghColores)
    ml_ColorLetra = lValue
End Property


Property Let EsNegrita(lValue As Boolean)
    ml_EsNegrita = lValue
End Property

Property Let LongitudLetra(lValue As String)
    ml_LongitudLetra = lValue
End Property

Property Let Mensaje(lValue As String)
    ml_Mensaje = lValue
End Property

Property Let Titulo(lValue As String)
    ml_Titulo = lValue
End Property

Private Sub btnAceptar_Click()
    Me.Visible = False
    mi_BotonPresionado = sghAceptar
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
    mi_BotonPresionado = sghCancelar
End Sub

Private Sub cmdCancelar_Click()
    Me.Visible = False
    mi_BotonPresionado = sghCancelar
End Sub

Private Sub Form_Load()
    Me.Caption = ml_Titulo
    '
    If ml_LongitudLetra = 0 Then
        Dim lnLineasMensaje As Integer
        lnLineasMensaje = 1
        If Len(ml_Mensaje) < 100 Then
           lblMensaje.Caption = Chr(13) & Chr(13) & Chr(13) & Chr(13) & ml_Mensaje
        ElseIf Len(ml_Mensaje) >= 100 And Len(ml_Mensaje) < 200 Then
           lblMensaje.Caption = Chr(13) & Chr(13) & Chr(13) & ml_Mensaje
        ElseIf Len(ml_Mensaje) >= 200 And Len(ml_Mensaje) < 300 Then
           lblMensaje.Caption = Chr(13) & Chr(13) & ml_Mensaje
        ElseIf Len(ml_Mensaje) >= 300 And Len(ml_Mensaje) < 400 Then
           lblMensaje.Caption = Chr(13) & ml_Mensaje
        ElseIf Len(ml_Mensaje) >= 400 Then
           lblMensaje.Caption = ml_Mensaje
        End If
    Else
        lblMensaje.Caption = ml_Mensaje
    End If
    '
    If ml_LongitudLetra > 0 Then
       lblMensaje.Font.Size = ml_LongitudLetra
    Else
       lblMensaje.Font.Size = 9
    End If
    If ml_EsNegrita = True Then
       lblMensaje.Font.Bold = True
    Else
       lblMensaje.Font.Bold = False
    End If
    Select Case ml_ColorLetra
    Case sghAzul
        lblMensaje.ForeColor = vbBlue
    Case sghBlanco
        lblMensaje.ForeColor = vbWhite
    Case sghNegro
        lblMensaje.ForeColor = vbBlack
    Case sghRojo
        lblMensaje.ForeColor = vbRed
    Case sghVerde
        lblMensaje.ForeColor = vbGreen
    Case Else
        lblMensaje.ForeColor = vbRed
    End Select
    If ml_UsaBotonesAceptarCancelar = True Then
       Me.btnCancelar.Visible = False
       Me.btnAceptar.Visible = True
       Me.cmdCancelar.Visible = True
    Else
       Me.btnCancelar.Visible = True
       Me.btnAceptar.Visible = False
       Me.cmdCancelar.Visible = False
    End If
End Sub



