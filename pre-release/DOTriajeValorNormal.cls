VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTriajeValorNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TriajeValorNormal
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdTriajeValorNormal As Long
Dim ml_EdadInicialEnDia As Long
Dim ml_EdadFinalEnDia As Long
Dim mc_ValorNormalMinimo As Currency
Dim mc_ValorNormalMaximo As Currency
Dim mc_ValorCoherenteMinimo As Currency
Dim mc_ValorCoherenteMaximo As Currency
Dim ml_IdTriajeVariable As Long
Dim ml_EstadoPaciente As Long
Dim ml_SexoPaciente As Long
Dim mda_FechaVigencia As Date

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdTriajeValorNormal(iValue As Long)
   ml_IdTriajeValorNormal = iValue
End Property
Property Get IdTriajeValorNormal() As Long
   IdTriajeValorNormal = ml_IdTriajeValorNormal
End Property
Property Let EdadInicialEnDia(iValue As Long)
   ml_EdadInicialEnDia = iValue
End Property
Property Get EdadInicialEnDia() As Long
   EdadInicialEnDia = ml_EdadInicialEnDia
End Property
Property Let EdadFinalEnDia(iValue As Long)
   ml_EdadFinalEnDia = iValue
End Property
Property Get EdadFinalEnDia() As Long
   EdadFinalEnDia = ml_EdadFinalEnDia
End Property
Property Let ValorNormalMinimo(cValue As Currency)
   mc_ValorNormalMinimo = cValue
End Property
Property Get ValorNormalMinimo() As Currency
   ValorNormalMinimo = mc_ValorNormalMinimo
End Property
Property Let ValorNormalMaximo(cValue As Currency)
   mc_ValorNormalMaximo = cValue
End Property
Property Get ValorNormalMaximo() As Currency
   ValorNormalMaximo = mc_ValorNormalMaximo
End Property
Property Let ValorCoherenteMinimo(cValue As Currency)
   mc_ValorCoherenteMinimo = cValue
End Property
Property Get ValorCoherenteMinimo() As Currency
   ValorCoherenteMinimo = mc_ValorCoherenteMinimo
End Property
Property Let ValorCoherenteMaximo(cValue As Currency)
   mc_ValorCoherenteMaximo = cValue
End Property
Property Get ValorCoherenteMaximo() As Currency
   ValorCoherenteMaximo = mc_ValorCoherenteMaximo
End Property
Property Let IdTriajeVariable(iValue As Long)
   ml_IdTriajeVariable = iValue
End Property
Property Get IdTriajeVariable() As Long
   IdTriajeVariable = ml_IdTriajeVariable
End Property
Property Let EstadoPaciente(iValue As Long)
   ml_EstadoPaciente = iValue
End Property
Property Get EstadoPaciente() As Long
   EstadoPaciente = ml_EstadoPaciente
End Property
Property Let SexoPaciente(iValue As Long)
   ml_SexoPaciente = iValue
End Property
Property Get SexoPaciente() As Long
   SexoPaciente = ml_SexoPaciente
End Property
Property Let FechaVigencia(daValue As Date)
   mda_FechaVigencia = daValue
End Property
Property Get FechaVigencia() As Date
   FechaVigencia = mda_FechaVigencia
End Property
