VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form ResumenDeCuentaDetalle 
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "ResumenDeCuentaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   30
      TabIndex        =   4
      Top             =   1950
      Width           =   8355
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ResumenDeCuentaDetalle.frx":000C
         DownPicture     =   "ResumenDeCuentaDetalle.frx":04D0
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
         Left            =   4253
         Picture         =   "ResumenDeCuentaDetalle.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimeFichaSIS 
         Caption         =   "Imp.Ficha SIS"
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
         Left            =   2813
         Picture         =   "ResumenDeCuentaDetalle.frx":0EA8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Prepare su Impresora con el FORMATO UNICO DE ATENCION (FUA SIS)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8370
      Begin Threed.SSOption optCabecera 
         Height          =   345
         Left            =   150
         TabIndex        =   1
         Top             =   390
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   609
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
         Caption         =   "Sólo Imprime datos de Cabecera"
         Value           =   -1
      End
      Begin Threed.SSOption optDetalle 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   855
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   609
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
         Caption         =   "Sólo Imprime datos de la Atención del Médico"
      End
      Begin Threed.SSOption optCabeceraDetalle 
         Height          =   345
         Left            =   150
         TabIndex        =   3
         Top             =   1320
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   609
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
         Caption         =   "Imprime datos de la Cabecera y Atención del Médico"
      End
   End
End
Attribute VB_Name = "ResumenDeCuentaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resume de cuenta detalle
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_OpcionDefault  As Integer
Dim mo_IdAtencion As Long

Property Let idAtencion(lValue As Long)
   mo_IdAtencion = lValue
End Property

Property Let OpcionDefault(lValue As Integer)
   mo_OpcionDefault = lValue
End Property

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub btnImprimeFichaSIS_Click()
    Dim oImprimeSIS As New RptHistoriaClinicaCE
    If Me.optCabecera.Value = True Then
        oImprimeSIS.ImprimeFormatoSIS mo_IdAtencion, 0, 1
    ElseIf Me.optDetalle.Value = True Then
        oImprimeSIS.ImprimeFormatoSIS mo_IdAtencion, 0, 2
    Else
        oImprimeSIS.ImprimeFormatoSIS mo_IdAtencion, 0, 3
    End If
    btnCancelar_Click
End Sub

Private Sub Form_Load()
    Select Case mo_OpcionDefault
    Case 1
        Me.optCabecera.Value = True
        Me.optDetalle.Value = False
        Me.optCabeceraDetalle.Value = False
    Case 2
        Me.optCabecera.Value = False
        Me.optDetalle.Value = True
        Me.optCabeceraDetalle.Value = False
    Case Else
        Me.optCabecera.Value = False
        Me.optDetalle.Value = False
        Me.optCabeceraDetalle.Value = True
    End Select
End Sub



