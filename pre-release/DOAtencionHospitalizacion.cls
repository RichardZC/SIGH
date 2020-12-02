VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionHospitalizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOAtencionesHospitalizacion
'        Autor: William Castro Grijalva
'        Fecha: 11/10/2004 09:27:29 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdCamaEgreso As Long
Dim ml_IdCamaIngreso As Long
Dim mb_TieneNecropsia As Boolean
Dim mb_HuboInfeccionIntraHospitalaria As Boolean
Dim ml_IdServicioEgreso As Long
Dim ml_IdTipoAlta As Long
Dim ml_IdCondicionAlta As Long
Dim ml_IdAtencion As Long
Dim ml_IdAtencionHospitalizacion As Long
Dim ms_HoraEgresoAdministrativo As String
Dim mda_FechaEgresoAdministrativo As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdCamaEgreso(lValue As Long)
   ml_IdCamaEgreso = lValue
End Property
Property Get IdCamaEgreso() As Long
   IdCamaEgreso = ml_IdCamaEgreso
End Property
Property Let IdCamaIngreso(lValue As Long)
   ml_IdCamaIngreso = lValue
End Property
Property Get IdCamaIngreso() As Long
   IdCamaIngreso = ml_IdCamaIngreso
End Property
Property Let TieneNecropsia(bValue As Boolean)
   mb_TieneNecropsia = bValue
End Property
Property Get TieneNecropsia() As Boolean
   TieneNecropsia = mb_TieneNecropsia
End Property
Property Let HuboInfeccionIntraHospitalaria(bValue As Boolean)
   mb_HuboInfeccionIntraHospitalaria = bValue
End Property
Property Get HuboInfeccionIntraHospitalaria() As Boolean
   HuboInfeccionIntraHospitalaria = mb_HuboInfeccionIntraHospitalaria
End Property
Property Let IdServicioEgreso(lValue As Long)
   ml_IdServicioEgreso = lValue
End Property
Property Get IdServicioEgreso() As Long
   IdServicioEgreso = ml_IdServicioEgreso
End Property
Property Let IdTipoAlta(lValue As Long)
   ml_IdTipoAlta = lValue
End Property
Property Get IdTipoAlta() As Long
   IdTipoAlta = ml_IdTipoAlta
End Property
Property Let IdCondicionAlta(lValue As Long)
   ml_IdCondicionAlta = lValue
End Property
Property Get IdCondicionAlta() As Long
   IdCondicionAlta = ml_IdCondicionAlta
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdAtencionHospitalizacion(lValue As Long)
   ml_IdAtencionHospitalizacion = lValue
End Property
Property Get IdAtencionHospitalizacion() As Long
   IdAtencionHospitalizacion = ml_IdAtencionHospitalizacion
End Property
Property Let HoraEgresoAdministrativo(sValue As String)
   ms_HoraEgresoAdministrativo = sValue
End Property
Property Get HoraEgresoAdministrativo() As String
   HoraEgresoAdministrativo = ms_HoraEgresoAdministrativo
End Property
Property Let FechaEgresoAdministrativo(daValue As Date)
   mda_FechaEgresoAdministrativo = daValue
End Property
Property Get FechaEgresoAdministrativo() As Date
   FechaEgresoAdministrativo = mda_FechaEgresoAdministrativo
End Property


