VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoImagMovimientoResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: ImagMovimientoResultados
'        Autor: Daniel Enrique Barrantes
'        Fecha: 23/10/2018 06:15:54 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idMovimiento As Long
Dim ml_idProductoCPT As Long
Dim ms_Resultado As String
Dim mda_ResultadoFecha As Date
Dim ms_ResultadoDNI As String
Dim ms_EquipoRuta As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idMovimiento(iValue As Long)
   ml_idMovimiento = iValue
End Property
Property Get idMovimiento() As Long
   idMovimiento = ml_idMovimiento
End Property
Property Let idProductoCPT(iValue As Long)
   ml_idProductoCPT = iValue
End Property
Property Get idProductoCPT() As Long
   idProductoCPT = ml_idProductoCPT
End Property
Property Let Resultado(sValue As String)
   ms_Resultado = sValue
End Property
Property Get Resultado() As String
   Resultado = ms_Resultado
End Property
Property Let ResultadoFecha(daValue As Date)
   mda_ResultadoFecha = daValue
End Property
Property Get ResultadoFecha() As Date
   ResultadoFecha = mda_ResultadoFecha
End Property
Property Let ResultadoDNI(sValue As String)
   ms_ResultadoDNI = sValue
End Property
Property Get ResultadoDNI() As String
   ResultadoDNI = ms_ResultadoDNI
End Property
Property Let EquipoRuta(sValue As String)
   ms_EquipoRuta = sValue
End Property
Property Get EquipoRuta() As String
   EquipoRuta = ms_EquipoRuta
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: ImagMovimientoResultados
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: ImagMovimientoResultados
'        Autor:   DAniel Barrantes B.
'        Fecha: 23/10/2018 06:15:55 p.m.
'------------------------------------------------------------------------------------
 

