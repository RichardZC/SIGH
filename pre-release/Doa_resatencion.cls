VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Doa_resatencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ResAtencion
'        Programado por: Cachay F
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ms_pers_IdResAtencion As String
Dim ms_pers_IdTipoDocumento As String
Dim ms_pers_ApePaterno As String
Dim ms_pers_ApeMaterno As String
Dim ms_pers_PriNombre As String
Dim ms_pers_OtrNombre As String
Dim ms_pers_IdTipoPersonalSalud As String
Dim ms_pers_Colegiatura As String
Dim ms_pers_IdEspecialidad As String
Dim ms_pers_NroEspecialidad As String
Dim ms_pers_IdEstado As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let pers_IdResAtencion(sValue As String)
   ms_pers_IdResAtencion = sValue
End Property
Property Get pers_IdResAtencion() As String
   pers_IdResAtencion = ms_pers_IdResAtencion
End Property
Property Let pers_IdTipoDocumento(sValue As String)
   ms_pers_IdTipoDocumento = sValue
End Property
Property Get pers_IdTipoDocumento() As String
   pers_IdTipoDocumento = ms_pers_IdTipoDocumento
End Property
Property Let pers_ApePaterno(sValue As String)
   ms_pers_ApePaterno = sValue
End Property
Property Get pers_ApePaterno() As String
   pers_ApePaterno = ms_pers_ApePaterno
End Property
Property Let pers_ApeMaterno(sValue As String)
   ms_pers_ApeMaterno = sValue
End Property
Property Get pers_ApeMaterno() As String
   pers_ApeMaterno = ms_pers_ApeMaterno
End Property
Property Let pers_PriNombre(sValue As String)
   ms_pers_PriNombre = sValue
End Property
Property Get pers_PriNombre() As String
   pers_PriNombre = ms_pers_PriNombre
End Property
Property Let pers_OtrNombre(sValue As String)
   ms_pers_OtrNombre = sValue
End Property
Property Get pers_OtrNombre() As String
   pers_OtrNombre = ms_pers_OtrNombre
End Property
Property Let pers_IdTipoPersonalSalud(sValue As String)
   ms_pers_IdTipoPersonalSalud = sValue
End Property
Property Get pers_IdTipoPersonalSalud() As String
   pers_IdTipoPersonalSalud = ms_pers_IdTipoPersonalSalud
End Property
Property Let pers_Colegiatura(sValue As String)
   ms_pers_Colegiatura = sValue
End Property
Property Get pers_Colegiatura() As String
   pers_Colegiatura = ms_pers_Colegiatura
End Property
Property Let pers_IdEspecialidad(sValue As String)
   ms_pers_IdEspecialidad = sValue
End Property
Property Get pers_IdEspecialidad() As String
   pers_IdEspecialidad = ms_pers_IdEspecialidad
End Property
Property Let pers_NroEspecialidad(sValue As String)
   ms_pers_NroEspecialidad = sValue
End Property
Property Get pers_NroEspecialidad() As String
   pers_NroEspecialidad = ms_pers_NroEspecialidad
End Property
Property Let pers_IdEstado(sValue As String)
   ms_pers_IdEstado = sValue
End Property
Property Get pers_IdEstado() As String
   pers_IdEstado = ms_pers_IdEstado
End Property

