VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPerinatalAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PerinatalAtencion
'        Programado por: Garay M
'        Fecha: Setiembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPerinatalAtencion As Long
Dim ml_IdPaciente As Long
Dim ml_idModulo As Long
Dim ml_GrafXedadEnMeses As Long
Dim ml_GrafYpercentilTE As Long
Dim ml_GrafYpercentilPT As Long
Dim ml_GrafYpercentilPE As Long
Dim mc_GrafYimc As Currency
Dim mda_FechaAtencion As Date
Dim ml_CredN As Long
Property Let CredN(lValue As Long)
   ml_CredN = lValue
End Property
Property Get CredN() As Long
   CredN = ml_CredN
End Property
Property Let FechaAtencion(daValue As Date)
   mda_FechaAtencion = daValue
End Property
Property Get FechaAtencion() As Date
   FechaAtencion = mda_FechaAtencion
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idPerinatalAtencion(iValue As Long)
   ml_idPerinatalAtencion = iValue
End Property
Property Get idPerinatalAtencion() As Long
   idPerinatalAtencion = ml_idPerinatalAtencion
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let idModulo(iValue As Long)
   ml_idModulo = iValue
End Property
Property Get idModulo() As Long
   idModulo = ml_idModulo
End Property
Property Let GrafXedadEnMeses(iValue As Long)
   ml_GrafXedadEnMeses = iValue
End Property
Property Get GrafXedadEnMeses() As Long
   GrafXedadEnMeses = ml_GrafXedadEnMeses
End Property
Property Let GrafYpercentilTE(iValue As Long)
   ml_GrafYpercentilTE = iValue
End Property
Property Get GrafYpercentilTE() As Long
   GrafYpercentilTE = ml_GrafYpercentilTE
End Property
Property Let GrafYpercentilPT(iValue As Long)
   ml_GrafYpercentilPT = iValue
End Property
Property Get GrafYpercentilPT() As Long
   GrafYpercentilPT = ml_GrafYpercentilPT
End Property
Property Let GrafYpercentilPE(iValue As Long)
   ml_GrafYpercentilPE = iValue
End Property
Property Get GrafYpercentilPE() As Long
   GrafYpercentilPE = ml_GrafYpercentilPE
End Property
Property Let GrafYimc(cValue As Currency)
   mc_GrafYimc = cValue
End Property
Property Get GrafYimc() As Currency
   GrafYimc = mc_GrafYimc
End Property

