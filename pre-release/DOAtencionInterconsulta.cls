VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionInterconsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesInterconsultas
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdDetalleProducto As Long
Dim ml_IdAtencion As Long
Dim ms_HoraSolicitud As String
Dim ms_HoraRealizacion As String
Dim mda_FechaSolicitud As Date
Dim mda_FechaRealizacion As Date
Dim ml_IdMedicoRealiza As Long
Dim ml_IdMedicoSolicita As Long
Dim ml_IdInterconsulta As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdDetalleProducto(lValue As Long)
   ml_IdDetalleProducto = lValue
End Property
Property Get IdDetalleProducto() As Long
   IdDetalleProducto = ml_IdDetalleProducto
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let HoraSolicitud(sValue As String)
   ms_HoraSolicitud = sValue
End Property
Property Get HoraSolicitud() As String
   HoraSolicitud = ms_HoraSolicitud
End Property
Property Let HoraRealizacion(sValue As String)
   ms_HoraRealizacion = sValue
End Property
Property Get HoraRealizacion() As String
   HoraRealizacion = ms_HoraRealizacion
End Property
Property Let FechaSolicitud(daValue As Date)
   mda_FechaSolicitud = daValue
End Property
Property Get FechaSolicitud() As Date
   FechaSolicitud = mda_FechaSolicitud
End Property
Property Let FechaRealizacion(daValue As Date)
   mda_FechaRealizacion = daValue
End Property
Property Get FechaRealizacion() As Date
   FechaRealizacion = mda_FechaRealizacion
End Property
Property Let IdMedicoRealiza(lValue As Long)
   ml_IdMedicoRealiza = lValue
End Property
Property Get IdMedicoRealiza() As Long
   IdMedicoRealiza = ml_IdMedicoRealiza
End Property
Property Let IdMedicoSolicita(lValue As Long)
   ml_IdMedicoSolicita = lValue
End Property
Property Get IdMedicoSolicita() As Long
   IdMedicoSolicita = ml_IdMedicoSolicita
End Property
Property Let IdInterconsulta(lValue As Long)
   ml_IdInterconsulta = lValue
End Property
Property Get IdInterconsulta() As Long
   IdInterconsulta = ml_IdInterconsulta
End Property


