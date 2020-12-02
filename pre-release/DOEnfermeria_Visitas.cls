VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEnfermeria_Visitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Enfermeria_visitas
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdVisita As Long
Dim mda_FechaHoraVisita As Date
Dim ml_IdServicio As Long
Dim ms_Observaciones As String
Dim ml_IdCama As Long
Dim ml_IdEmpleadoEnfermera As Long
Dim mb_IngresoValorizacion As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdVisita(iValue As Long)
   ml_IdVisita = iValue
End Property
Property Get IdVisita() As Long
   IdVisita = ml_IdVisita
End Property
Property Let FechaHoraVisita(daValue As Date)
   mda_FechaHoraVisita = daValue
End Property
Property Get FechaHoraVisita() As Date
   FechaHoraVisita = mda_FechaHoraVisita
End Property
Property Let IdServicio(iValue As Long)
   ml_IdServicio = iValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let Observaciones(sValue As String)
   ms_Observaciones = sValue
End Property
Property Get Observaciones() As String
   Observaciones = ms_Observaciones
End Property
Property Let IdCama(iValue As Long)
   ml_IdCama = iValue
End Property
Property Get IdCama() As Long
   IdCama = ml_IdCama
End Property
Property Let IdEmpleadoEnfermera(iValue As Long)
   ml_IdEmpleadoEnfermera = iValue
End Property
Property Get IdEmpleadoEnfermera() As Long
   IdEmpleadoEnfermera = ml_IdEmpleadoEnfermera
End Property
Property Let IngresoValorizacion(bValue As Boolean)
   mb_IngresoValorizacion = bValue
End Property
Property Get IngresoValorizacion() As Boolean
   IngresoValorizacion = mb_IngresoValorizacion
End Property

