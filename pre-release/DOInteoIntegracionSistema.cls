VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOInteoIntegracionSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla InteoIntegracionSistema
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdIntegracionSistema As Long
Dim ml_IdTipoSistema As Long
Dim ml_IdProveedorSistema As Long
Dim ms_NombreUsuario As String
Dim ms_ClaveUsuario As String
Dim mb_EsProveedorActual As Boolean
Dim mb_EsActivo As Boolean
Dim mda_FechaCrea As Date
Dim mda_FechaEdita As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdIntegracionSistema(iValue As Long)
   ml_IdIntegracionSistema = iValue
End Property
Property Get IdIntegracionSistema() As Long
   IdIntegracionSistema = ml_IdIntegracionSistema
End Property
Property Let IdTipoSistema(iValue As Long)
   ml_IdTipoSistema = iValue
End Property
Property Get IdTipoSistema() As Long
   IdTipoSistema = ml_IdTipoSistema
End Property
Property Let IdProveedorSistema(iValue As Long)
   ml_IdProveedorSistema = iValue
End Property
Property Get IdProveedorSistema() As Long
   IdProveedorSistema = ml_IdProveedorSistema
End Property
Property Let NombreUsuario(sValue As String)
   ms_NombreUsuario = sValue
End Property
Property Get NombreUsuario() As String
   NombreUsuario = ms_NombreUsuario
End Property
Property Let ClaveUsuario(sValue As String)
   ms_ClaveUsuario = sValue
End Property
Property Get ClaveUsuario() As String
   ClaveUsuario = ms_ClaveUsuario
End Property
Property Let EsProveedorActual(bValue As Boolean)
   mb_EsProveedorActual = bValue
End Property
Property Get EsProveedorActual() As Boolean
   EsProveedorActual = mb_EsProveedorActual
End Property
Property Let EsActivo(bValue As Boolean)
   mb_EsActivo = bValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = mb_EsActivo
End Property
Property Let FechaCrea(daValue As Date)
   mda_FechaCrea = daValue
End Property
Property Get FechaCrea() As Date
   FechaCrea = mda_FechaCrea
End Property
Property Let FechaEdita(daValue As Date)
   mda_FechaEdita = daValue
End Property
Property Get FechaEdita() As Date
   FechaEdita = mda_FechaEdita
End Property
