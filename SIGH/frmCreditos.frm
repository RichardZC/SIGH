VERSION 5.00
Begin VB.Form Creditos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos"
   ClientHeight    =   3720
   ClientLeft      =   1125
   ClientTop       =   4545
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   """El único bien que se incrementa cuando se comparte es el conocimiento"""
      ForeColor       =   &H00B9553C&
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   3165
      Width           =   4035
   End
   Begin VB.Shape shTitulo 
      BorderColor     =   &H00B9553C&
      Height          =   3090
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image imgCredito 
      Height          =   3075
      Left            =   75
      Picture         =   "frmCreditos.frx":0000
      Top             =   60
      Width           =   4065
   End
End
Attribute VB_Name = "Creditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
