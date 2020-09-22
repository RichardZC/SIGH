VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoDatosTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mc_Presion As String
Dim mc_Peso As String
Dim mc_Talla As String
Dim mc_Temperatura As String
Dim ml_Pulso As Long
Dim ml_FrecuenciaRespiratoria As Long
Dim ml_FrecuenciaCardiaca As Long
Dim ml_Pulso As Currency

