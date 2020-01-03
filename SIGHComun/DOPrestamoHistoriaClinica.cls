VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPrestamoHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOPrestamosHistoriaClinica
'        Autor: William Castro Grijalva
'        Fecha: 15/01/2005 03:29:02 p.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_Auditoria As Long
Dim ml_IdMotivo As Long
Dim mda_FechaPrestamoRequerida As Date
Dim ms_HoraPrestamoRequerida As String
Dim mda_FechaPrestamoReal As Date
Dim ml_IdPrestamo As Long
Dim mda_FechaSolicitud As Date
Dim ms_HoraSolicitud As String
Dim ml_IdEstadoPrestamo As Long
Dim ml_IdPaciente As Long
Dim ml_IdEnvio As Long
Dim ms_Observacion As String
Dim ml_IdServicio As Long
Dim ms_HoraPrestamoReal As String
Dim ms_HoraDevolucion As String
Dim ml_NroFolios As Long
Dim mda_FechaDevolucion As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdMotivo(lValue As Long)
   ml_IdMotivo = lValue
End Property
Property Get IdMotivo() As Long
   IdMotivo = ml_IdMotivo
End Property
Property Let FechaPrestamoRequerida(daValue As Date)
   mda_FechaPrestamoRequerida = daValue
End Property
Property Get FechaPrestamoRequerida() As Date
   FechaPrestamoRequerida = mda_FechaPrestamoRequerida
End Property
Property Let HoraPrestamoRequerida(sValue As String)
   ms_HoraPrestamoRequerida = sValue
End Property
Property Get HoraPrestamoRequerida() As String
   HoraPrestamoRequerida = ms_HoraPrestamoRequerida
End Property
Property Let FechaPrestamoReal(daValue As Date)
   mda_FechaPrestamoReal = daValue
End Property
Property Get FechaPrestamoReal() As Date
   FechaPrestamoReal = mda_FechaPrestamoReal
End Property
Property Let IdPrestamo(lValue As Long)
   ml_IdPrestamo = lValue
End Property
Property Get IdPrestamo() As Long
   IdPrestamo = ml_IdPrestamo
End Property
Property Let FechaSolicitud(daValue As Date)
   mda_FechaSolicitud = daValue
End Property
Property Get FechaSolicitud() As Date
   FechaSolicitud = mda_FechaSolicitud
End Property
Property Let HoraSolicitud(sValue As String)
   ms_HoraSolicitud = sValue
End Property
Property Get HoraSolicitud() As String
   HoraSolicitud = ms_HoraSolicitud
End Property
Property Let IdEstadoPrestamo(lValue As Long)
   ml_IdEstadoPrestamo = lValue
End Property
Property Get IdEstadoPrestamo() As Long
   IdEstadoPrestamo = ml_IdEstadoPrestamo
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdEnvio(lValue As Long)
   ml_IdEnvio = lValue
End Property
Property Get IdEnvio() As Long
   IdEnvio = ml_IdEnvio
End Property
Property Let Observacion(sValue As String)
   ms_Observacion = sValue
End Property
Property Get Observacion() As String
   Observacion = ms_Observacion
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let HoraPrestamoReal(sValue As String)
   ms_HoraPrestamoReal = sValue
End Property
Property Get HoraPrestamoReal() As String
   HoraPrestamoReal = ms_HoraPrestamoReal
End Property
Property Let HoraDevolucion(sValue As String)
   ms_HoraDevolucion = sValue
End Property
Property Get HoraDevolucion() As String
   HoraDevolucion = ms_HoraDevolucion
End Property
Property Let NroFolios(lValue As Long)
   ml_NroFolios = lValue
End Property
Property Get NroFolios() As Long
   NroFolios = ml_NroFolios
End Property
Property Let FechaDevolucion(daValue As Date)
   mda_FechaDevolucion = daValue
End Property
Property Get FechaDevolucion() As Date
   FechaDevolucion = mda_FechaDevolucion
End Property


