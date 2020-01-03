VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTipoFinanciador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TipoFinanciador
'        Programado por: Garay M
'        Fecha: Marzo 2015
'
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idTipoFinanciador As Long
Dim ms_nombre As String
Dim ms_denominacion As String
Dim ms_codigo As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idTipoFinanciador(iValue As Long)
   ml_idTipoFinanciador = iValue
End Property
Property Get idTipoFinanciador() As Long
   idTipoFinanciador = ml_idTipoFinanciador
End Property
Property Let nombre(sValue As String)
   ms_nombre = sValue
End Property
Property Get nombre() As String
   nombre = ms_nombre
End Property
Property Let denominacion(sValue As String)
   ms_denominacion = sValue
End Property
Property Get denominacion() As String
   denominacion = ms_denominacion
End Property
Property Let codigo(sValue As String)
   ms_codigo = sValue
End Property
Property Get codigo() As String
   codigo = ms_codigo
End Property
