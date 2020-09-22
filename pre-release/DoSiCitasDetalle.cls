VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSiCitasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SiCitasDetalle
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/11/2018 09:19:28 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idCitaSI As Long
Dim ml_idProductoCpt As Long
Dim ml_cantidad As Long
Dim mc_precio As Currency
Dim ms_Observaciones As String
Dim ms_llaveTicket As String
Property Let llaveTicket(sValue As String)
   ms_llaveTicket = sValue
End Property
Property Get llaveTicket() As String
   llaveTicket = ms_llaveTicket
End Property

Property Let Observaciones(sValue As String)
   ms_Observaciones = sValue
End Property
Property Get Observaciones() As String
   Observaciones = ms_Observaciones
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idCitaSI(iValue As Long)
   ml_idCitaSI = iValue
End Property
Property Get idCitaSI() As Long
   idCitaSI = ml_idCitaSI
End Property
Property Let idProductoCpt(iValue As Long)
   ml_idProductoCpt = iValue
End Property
Property Get idProductoCpt() As Long
   idProductoCpt = ml_idProductoCpt
End Property
Property Let cantidad(iValue As Long)
   ml_cantidad = iValue
End Property
Property Get cantidad() As Long
   cantidad = ml_cantidad
End Property
Property Let precio(cValue As Currency)
   mc_precio = cValue
End Property
Property Get precio() As Currency
   precio = mc_precio
End Property


'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: SiCitasDetalle
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SiCitasDetalle
'        Autor:   DAniel Barrantes B.
'        Fecha: 13/11/2018 09:19:28 a.m.
'------------------------------------------------------------------------------------
 

