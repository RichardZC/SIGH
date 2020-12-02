VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesEmergencia
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdTipoAgenteAGAN As Long
Dim ml_IdGrupoOcupacionalALAB As Long
Dim ml_IdPosicionLesionadoALAB As Long
Dim ml_IdUbicacionLesionado As Long
Dim ml_IdTipoTransporte As Long
Dim ml_IdTipoVehiculo As Long
Dim ml_IdClaseAccidente As Long
Dim ml_IdRelacionAgresorVictima As Long
Dim ml_IdSeguridad As Long
Dim ml_IdTipoEvento As Long
Dim ml_IdLugarEvento As Long
Dim ml_IdCausaExternaMorbilidad As Long
Dim ml_IdAtencion As Long
Dim ml_IdAtencionEmergencia As Long
Dim ml_comoLlego As Long
Dim ml_tipoAtencion As Long
Dim ml_idEstadoLlegada As Long

Property Let idEstadoLlegada(lValue As Long)
   ml_idEstadoLlegada = lValue
End Property
Property Get idEstadoLlegada() As Long
   idEstadoLlegada = ml_idEstadoLlegada
End Property



Property Let tipoAtencion(lValue As Long)
   ml_tipoAtencion = lValue
End Property
Property Get tipoAtencion() As Long
   tipoAtencion = ml_tipoAtencion
End Property

Property Let comoLlego(lValue As Long)
   ml_comoLlego = lValue
End Property
Property Get comoLlego() As Long
   comoLlego = ml_comoLlego
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdTipoAgenteAGAN(lValue As Long)
   ml_IdTipoAgenteAGAN = lValue
End Property
Property Get IdTipoAgenteAGAN() As Long
   IdTipoAgenteAGAN = ml_IdTipoAgenteAGAN
End Property
Property Let IdGrupoOcupacionalALAB(lValue As Long)
   ml_IdGrupoOcupacionalALAB = lValue
End Property
Property Get IdGrupoOcupacionalALAB() As Long
   IdGrupoOcupacionalALAB = ml_IdGrupoOcupacionalALAB
End Property
Property Let IdPosicionLesionadoALAB(lValue As Long)
   ml_IdPosicionLesionadoALAB = lValue
End Property
Property Get IdPosicionLesionadoALAB() As Long
   IdPosicionLesionadoALAB = ml_IdPosicionLesionadoALAB
End Property
Property Let IdUbicacionLesionado(lValue As Long)
   ml_IdUbicacionLesionado = lValue
End Property
Property Get IdUbicacionLesionado() As Long
   IdUbicacionLesionado = ml_IdUbicacionLesionado
End Property
Property Let IdTipoTransporte(lValue As Long)
   ml_IdTipoTransporte = lValue
End Property
Property Get IdTipoTransporte() As Long
   IdTipoTransporte = ml_IdTipoTransporte
End Property
Property Let IdTipoVehiculo(lValue As Long)
   ml_IdTipoVehiculo = lValue
End Property
Property Get IdTipoVehiculo() As Long
   IdTipoVehiculo = ml_IdTipoVehiculo
End Property
Property Let IdClaseAccidente(lValue As Long)
   ml_IdClaseAccidente = lValue
End Property
Property Get IdClaseAccidente() As Long
   IdClaseAccidente = ml_IdClaseAccidente
End Property
Property Let IdRelacionAgresorVictima(lValue As Long)
   ml_IdRelacionAgresorVictima = lValue
End Property
Property Get IdRelacionAgresorVictima() As Long
   IdRelacionAgresorVictima = ml_IdRelacionAgresorVictima
End Property
Property Let IdSeguridad(lValue As Long)
   ml_IdSeguridad = lValue
End Property
Property Get IdSeguridad() As Long
   IdSeguridad = ml_IdSeguridad
End Property
Property Let IdTipoEvento(lValue As Long)
   ml_IdTipoEvento = lValue
End Property
Property Get IdTipoEvento() As Long
   IdTipoEvento = ml_IdTipoEvento
End Property
Property Let IdLugarEvento(lValue As Long)
   ml_IdLugarEvento = lValue
End Property
Property Get IdLugarEvento() As Long
   IdLugarEvento = ml_IdLugarEvento
End Property
Property Let IdCausaExternaMorbilidad(lValue As Long)
   ml_IdCausaExternaMorbilidad = lValue
End Property
Property Get IdCausaExternaMorbilidad() As Long
   IdCausaExternaMorbilidad = ml_IdCausaExternaMorbilidad
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdAtencionEmergencia(lValue As Long)
   ml_IdAtencionEmergencia = lValue
End Property
Property Get IdAtencionEmergencia() As Long
   IdAtencionEmergencia = ml_IdAtencionEmergencia
End Property


