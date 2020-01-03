VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Doa_presentaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: a_presentaciones
'        Autor: Daniel Enrique Barrantes
'        Fecha: 06/12/2013 12:52:27 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ms_tpre_IdPresentacion As String
Dim ms_tpre_Descripcion As String
Dim ms_tpre_Abreviatura As String
Dim ml_tpre_TopeMinimo As Long
Dim ml_tpre_TopeNoHosp As Long
Dim ml_tpre_TopeHosp As Long
Dim ms_tpre_IdEstado As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let tpre_IdPresentacion(sValue As String)
   ms_tpre_IdPresentacion = sValue
End Property
Property Get tpre_IdPresentacion() As String
   tpre_IdPresentacion = ms_tpre_IdPresentacion
End Property
Property Let tpre_Descripcion(sValue As String)
   ms_tpre_Descripcion = sValue
End Property
Property Get tpre_Descripcion() As String
   tpre_Descripcion = ms_tpre_Descripcion
End Property
Property Let tpre_Abreviatura(sValue As String)
   ms_tpre_Abreviatura = sValue
End Property
Property Get tpre_Abreviatura() As String
   tpre_Abreviatura = ms_tpre_Abreviatura
End Property
Property Let tpre_TopeMinimo(iValue As Long)
   ml_tpre_TopeMinimo = iValue
End Property
Property Get tpre_TopeMinimo() As Long
   tpre_TopeMinimo = ml_tpre_TopeMinimo
End Property
Property Let tpre_TopeNoHosp(iValue As Long)
   ml_tpre_TopeNoHosp = iValue
End Property
Property Get tpre_TopeNoHosp() As Long
   tpre_TopeNoHosp = ml_tpre_TopeNoHosp
End Property
Property Let tpre_TopeHosp(iValue As Long)
   ml_tpre_TopeHosp = iValue
End Property
Property Get tpre_TopeHosp() As Long
   tpre_TopeHosp = ml_tpre_TopeHosp
End Property
Property Let tpre_IdEstado(sValue As String)
   ms_tpre_IdEstado = sValue
End Property
Property Get tpre_IdEstado() As String
   tpre_IdEstado = ms_tpre_IdEstado
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: a_presentaciones
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: a_presentaciones
'        Autor:   DAniel Barrantes B.
'        Fecha: 06/12/2013 12:52:27 p.m.
'------------------------------------------------------------------------------------
 


