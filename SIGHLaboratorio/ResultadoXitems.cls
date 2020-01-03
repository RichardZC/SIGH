VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResultadoXitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Resultado de Resultado por Items
'        Programado por: Madrid S
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_IdMovimiento As New frmResultadoXitems
Dim ml_idUsuario As Long
Dim ml_idOrden As Long
Dim ml_idProductoCpt As Long
Dim ml_NoMuestraBotonGrabar As Boolean
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ml_idTipoSexo As Long
Dim ml_FechaNacimiento As Date

Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property
Property Let FechaNacimiento(lValue As Date)
    ml_FechaNacimiento = lValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let NoMuestraBotonGrabar(lValue As Boolean)
   ml_NoMuestraBotonGrabar = lValue
End Property

Property Let idProductoCPT(lValue As Long)
   ml_idProductoCpt = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property

Sub MostrarFormulario()
  mo_IdMovimiento.idOrden = ml_idOrden
  mo_IdMovimiento.idProductoCPT = ml_idProductoCpt
  mo_IdMovimiento.idUsuario = ml_idUsuario
  mo_IdMovimiento.NoMuestraBotonGrabar = ml_NoMuestraBotonGrabar
  mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
  mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
  mo_IdMovimiento.idTipoSexo = ml_idTipoSexo
  mo_IdMovimiento.FechaNacimiento = ml_FechaNacimiento
  mo_IdMovimiento.Show vbModal
  Set mo_IdMovimiento = Nothing
End Sub

