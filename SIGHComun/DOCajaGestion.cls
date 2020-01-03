VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CajaGestion
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ms_TotalCobrado As String
Dim mda_FechaCierre As Date
Dim ml_idTurno As Long
Dim ml_IdCaja As Long
Dim ml_IdCajero As Long
Dim ms_EstadoLote As String
Dim mda_FechaApertura As Date
Dim ml_IdGestionCaja As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let TotalCobrado(sValue As String)
   ms_TotalCobrado = sValue
End Property
Property Get TotalCobrado() As String
   TotalCobrado = ms_TotalCobrado
End Property
Property Let FechaCierre(daValue As Date)
   mda_FechaCierre = daValue
End Property
Property Get FechaCierre() As Date
   FechaCierre = mda_FechaCierre
End Property
Property Let idTurno(lValue As Long)
   ml_idTurno = lValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property
Property Let IdCaja(lValue As Long)
   ml_IdCaja = lValue
End Property
Property Get IdCaja() As Long
   IdCaja = ml_IdCaja
End Property
Property Let IdCajero(lValue As Long)
   ml_IdCajero = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property
Property Let EstadoLote(sValue As String)
   ms_EstadoLote = sValue
End Property
Property Get EstadoLote() As String
   EstadoLote = ms_EstadoLote
End Property
Property Let FechaApertura(daValue As Date)
   mda_FechaApertura = daValue
End Property
Property Get FechaApertura() As Date
   FechaApertura = mda_FechaApertura
End Property
Property Let IdGestionCaja(lValue As Long)
   ml_IdGestionCaja = lValue
End Property
Property Get IdGestionCaja() As Long
   IdGestionCaja = ml_IdGestionCaja
End Property


