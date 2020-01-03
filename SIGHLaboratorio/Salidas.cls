VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Salidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para salidas de insumos
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_IdMovimiento As New LabSalidas
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdMovimiento As Long
Dim ml_IdPuntoCarga As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
  mo_IdMovimiento.Opcion = mi_Opcion
  mo_IdMovimiento.idUsuario = ml_idUsuario
  mo_IdMovimiento.IdMovimiento = ml_IdMovimiento
  mo_IdMovimiento.IdPuntoCarga = ml_IdPuntoCarga
  mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
  mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
  mo_IdMovimiento.Show 1
  Set mo_IdMovimiento = Nothing
End Sub

Property Get Opcion() As sghOpciones
  Opcion = mi_Opcion
End Property

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Property Let Opcion(lValue As sghOpciones)
  mi_Opcion = lValue
End Property

Property Let IdMovimiento(lValue As Long)
  ml_IdMovimiento = lValue
End Property

Property Let IdPuntoCarga(lValue As Long)
  ml_IdPuntoCarga = lValue
End Property
