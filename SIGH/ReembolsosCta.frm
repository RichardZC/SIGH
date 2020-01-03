VERSION 5.00
Begin VB.Form ReembolsosCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   Icon            =   "ReembolsosCta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   90
      TabIndex        =   1
      Top             =   8430
      Width           =   12600
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ReembolsosCta.frx":0CCA
         DownPicture     =   "ReembolsosCta.frx":118E
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
         Left            =   6427
         Picture         =   "ReembolsosCta.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin SISGalenPlus.ucEstadoCuenta ucEstadoCuenta1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12675
      _extentx        =   22357
      _extenty        =   13785
   End
End
Attribute VB_Name = "ReembolsosCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reembolsos, consulta cuenta
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Dim ml_idCuentaAtencion As Long
Dim ml_GrabaConsumosConsolidados  As Boolean
Property Let GrabaConsumosConsolidados(lValue As Boolean)
   ml_GrabaConsumosConsolidados = lValue
End Property

Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property


Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Private Sub Form_Activate()
    If ml_GrabaConsumosConsolidados = True Then
       Me.Visible = False
    End If
End Sub

Private Sub Form_Load()
    ucEstadoCuenta1.Inicializar
    ucEstadoCuenta1.GrabaConsumosConsolidados = ml_GrabaConsumosConsolidados
    ucEstadoCuenta1.ConsultaDetalleCuenta (ml_idCuentaAtencion)

End Sub

