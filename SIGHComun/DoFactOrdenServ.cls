VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFactOrdenServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 '------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactOrdenServicio
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdOrden As Long
Dim ml_idPuntoCarga As Long
Dim ml_idPaciente As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idFuenteFinanciamiento As Long
Dim mda_FechaCreacion As Date
Dim ml_idUsuario As Long
Dim mda_FechaDespacho As Date
Dim ml_IdUsuarioDespacho As Long
Dim ml_IdEstadoFacturacion As Long
Dim ml_IdServicioPaciente As Long
Dim mda_FechaHoraRealizaCpt As Date
'SCCQ 23-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
'RHA 17/11/2020 Cambio 46 Inicio
'Dim ml_CodPres As String
'Property Get CodPres() As String
'   CodPres = ml_CodPres
'End Property
'Property Let CodPres(sValue As String)
'   ml_CodPres = sValue
'End Property
'RHA 17/11/2020 Cambio 46 Fin
'SCCQ 23-04-2021 Cambio 64 Fin (Reversion Cambio 46)


Property Let FechaHoraRealizaCpt(daValue As Date)
   mda_FechaHoraRealizaCpt = daValue
End Property
Property Get FechaHoraRealizaCpt() As Date
   FechaHoraRealizaCpt = mda_FechaHoraRealizaCpt
End Property

Property Let IdServicioPaciente(lValue As Long)
   ml_IdServicioPaciente = lValue
End Property
Property Get IdServicioPaciente() As Long
   IdServicioPaciente = ml_IdServicioPaciente
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property
Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property
Property Let idPuntoCarga(iValue As Long)
   ml_idPuntoCarga = iValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property
Property Let idPaciente(iValue As Long)
   ml_idPaciente = iValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property

Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let idFuenteFinanciamiento(iValue As Long)
   ml_idFuenteFinanciamiento = iValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdEstadoFacturacion(iValue As Long)
   ml_IdEstadoFacturacion = iValue
End Property
Property Get IdEstadoFacturacion() As Long
   IdEstadoFacturacion = ml_IdEstadoFacturacion
End Property
Property Let FechaDespacho(daValue As Date)
   mda_FechaDespacho = daValue
End Property
Property Get FechaDespacho() As Date
   FechaDespacho = mda_FechaDespacho
End Property
Property Let IdUsuarioDespacho(iValue As Long)
   ml_IdUsuarioDespacho = iValue
End Property
Property Get IdUsuarioDespacho() As Long
   IdUsuarioDespacho = ml_IdUsuarioDespacho
End Property


 

