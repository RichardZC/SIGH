VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoInteroperaEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InteroperaEquipos
'        Autor: Daniel Enrique Barrantes
'        Fecha: 23/10/2018 06:15:40 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ms_Codigo As String
Dim ms_Marca As String
Dim ms_Modelo As String
Dim ms_Tipo As String
Dim ms_Ruta As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let Marca(sValue As String)
   ms_Marca = sValue
End Property
Property Get Marca() As String
   Marca = ms_Marca
End Property
Property Let Modelo(sValue As String)
   ms_Modelo = sValue
End Property
Property Get Modelo() As String
   Modelo = ms_Modelo
End Property
Property Let Tipo(sValue As String)
   ms_Tipo = sValue
End Property
Property Get Tipo() As String
   Tipo = ms_Tipo
End Property
Property Let Ruta(sValue As String)
   ms_Ruta = sValue
End Property
Property Get Ruta() As String
   Ruta = ms_Ruta
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: InteroperaEquipos
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: InteroperaEquipos
'        Autor:   DAniel Barrantes B.
'        Fecha: 23/10/2018 06:15:41 a.m.
'------------------------------------------------------------------------------------
 

