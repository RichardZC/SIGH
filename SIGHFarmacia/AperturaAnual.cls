VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AperturaAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Apertura Anual
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Formulario As New HerrAperturaAnual
Dim mo_lcNombrePc  As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim ml_idUsuario As Long

Property Let idUsuario(lIdValue As Long)
    ml_idUsuario = lIdValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Sub MostrarFormulario()
    mo_Formulario.lcNombrePc = mo_lcNombrePc
    mo_Formulario.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_Formulario.lcNombrePc = mo_lcNombrePc
    mo_Formulario.idUsuario = ml_idUsuario
    mo_Formulario.Show 1
    Set mo_Formulario = Nothing
End Sub
