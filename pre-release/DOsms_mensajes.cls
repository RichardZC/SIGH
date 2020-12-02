VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOsms_mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: sms_mensajes
'        Autor: Daniel Enrique Barrantes
'        Fecha: 07/06/2018 10:07:46 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idMensaje As Long
Dim ms_celular As String
Dim ml_idCuentaAtencion As Long
Dim ms_Mensaje As String
Dim ml_idPaciente As Long
Dim mda_FechaEnvio As Date
Dim mda_FechaCreacion As Date
Dim ms_opcion As String
Dim ml_idEstadoEnvio As Long

Property Let IdMensaje(lValue As Long)
   ml_idMensaje = lValue
End Property
Property Get IdMensaje() As Long
   IdMensaje = ml_idMensaje
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let celular(sValue As String)
   ms_celular = sValue
End Property
Property Get celular() As String
   celular = ms_celular
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let Mensaje(sValue As String)
   ms_Mensaje = sValue
End Property
Property Get Mensaje() As String
   Mensaje = ms_Mensaje
End Property
Property Let idPaciente(iValue As Long)
   ml_idPaciente = iValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property
Property Let FechaEnvio(daValue As Date)
   mda_FechaEnvio = daValue
End Property
Property Get FechaEnvio() As Date
   FechaEnvio = mda_FechaEnvio
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let opcion(sValue As String)
   ms_opcion = sValue
End Property
Property Get opcion() As String
   opcion = ms_opcion
End Property
Property Let idEstadoEnvio(iValue As Long)
   ml_idEstadoEnvio = iValue
End Property
Property Get idEstadoEnvio() As Long
   idEstadoEnvio = ml_idEstadoEnvio
End Property

