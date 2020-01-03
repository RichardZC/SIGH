VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD Ini 01/06/2005 [Todo el archivo]
'MZD02 Ini 04/07/2005
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaLote
'        Autor: William Castro Grijalva
'        Fecha: 02/06/2005 08:51:58
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdCajero As Long
Dim mc_SaldoInicialDolares As Currency
Dim mc_SaldoInicialSoles As Currency
Dim ms_EstadoLote As String
Dim mda_Fecha As Date
Dim ml_IdLote As Long
Dim ml_IdCaja As Long
Dim ml_IdTurno As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdCajero(lValue As Long)
   ml_IdCajero = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property
Property Let SaldoInicialDolares(cValue As Currency)
   mc_SaldoInicialDolares = cValue
End Property
Property Get SaldoInicialDolares() As Currency
   SaldoInicialDolares = mc_SaldoInicialDolares
End Property
Property Let SaldoInicialSoles(cValue As Currency)
   mc_SaldoInicialSoles = cValue
End Property
Property Get SaldoInicialSoles() As Currency
   SaldoInicialSoles = mc_SaldoInicialSoles
End Property
Property Let EstadoLote(sValue As String)
   ms_EstadoLote = sValue
End Property
Property Get EstadoLote() As String
   EstadoLote = ms_EstadoLote
End Property
Property Let Fecha(daValue As Date)
   mda_Fecha = daValue
End Property
Property Get Fecha() As Date
   Fecha = mda_Fecha
End Property
Property Let IdLote(lValue As Long)
   ml_IdLote = lValue
End Property
Property Get IdLote() As Long
   IdLote = ml_IdLote
End Property
Property Let IdCaja(lValue As Long)
   ml_IdCaja = lValue
End Property
Property Get IdCaja() As Long
   IdCaja = ml_IdCaja
End Property
Property Let IdTurno(lValue As Long)
   ml_IdTurno = lValue
End Property
Property Get IdTurno() As Long
   IdTurno = ml_IdTurno
End Property






