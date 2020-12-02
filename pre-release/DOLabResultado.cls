VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "\DOLabResultado"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOLabResultado
'        Autor: Adams BONILLA MAGALLANES
'        Fecha: 25/06/2009 11:24:47 a.m.
'        Auspiciado por: Hospital Regional Ayacucho
'        Control De Cambios: 
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idAnalisis as Long
Dim ml_idOrden as Long
Dim ms_resultadoAnalisis as String
Dim ms_observacionResultado as String
Dim ml_idUsuario as Long

Property Let IdUsuarioAuditoria(lValue as Long)
   ml_IdUsuarioAuditoria = lValue
End Property

Property Get IdUsuarioAuditoria() as Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idAnalisis(iValue as Long)
   ml_idAnalisis = iValue
End Property

Property Get idAnalisis() as Long
   idAnalisis = ml_idAnalisis
End Property

Property Let idOrden(iValue as Long)
   ml_idOrden = iValue
End Property

Property Get idOrden() as Long
   idOrden = ml_idOrden
End Property

Property Let resultadoAnalisis(sValue as String)
   ms_resultadoAnalisis = sValue
End Property

Property Get resultadoAnalisis() as String
   resultadoAnalisis = ms_resultadoAnalisis
End Property

Property Let observacionResultado(sValue as String)
   ms_observacionResultado = sValue
End Property

Property Get observacionResultado() as String
   observacionResultado = ms_observacionResultado
End Property

Property Let idUsuario(iValue as Long)
   ml_idUsuario = iValue
End Property

Property Get idUsuario() as Long
   idUsuario = ml_idUsuario
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: DOLabResultado
'        Autor:   Adams BONILLA MAGALLANES  
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: DOLabResultado
'        Autor:   Adams BONILLA MAGALLANES 
'        Fecha: 25/06/2009 11:24:47 a.m.
'------------------------------------------------------------------------------------
 
