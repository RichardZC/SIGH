VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOProcedimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Procedimientos
'        Programado por: Castro W
'        Fecha: Diciembre 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_EdadMinDias As Long
Dim ml_IdProducto As Long
Dim ml_idTipoSexo As Long
Dim ms_DescripcionOPCS As String
Dim ms_CodigoOPCS As String
Dim ml_EdadMaxDias As Long
Dim mb_Restriccion As Boolean
Dim ms_CodigoCPT2004 As String
Dim ms_CodigoCPT99 As String
Dim ms_Descripcion As String
Dim ml_IdProcedimiento As Long
Dim mi_IdDiferenciacion As Integer 'ZCC20060322 (para diferencias los procedimientos para ayuda de diagnostico)
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let EdadMinDias(lValue As Long)
   ml_EdadMinDias = lValue
End Property
Property Get EdadMinDias() As Long
   EdadMinDias = ml_EdadMinDias
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property
Property Let DescripcionOPCS(sValue As String)
   ms_DescripcionOPCS = sValue
End Property
Property Get DescripcionOPCS() As String
   DescripcionOPCS = ms_DescripcionOPCS
End Property
Property Let CodigoOPCS(sValue As String)
   ms_CodigoOPCS = sValue
End Property
Property Get CodigoOPCS() As String
   CodigoOPCS = ms_CodigoOPCS
End Property
Property Let EdadMaxDias(lValue As Long)
   ml_EdadMaxDias = lValue
End Property
Property Get EdadMaxDias() As Long
   EdadMaxDias = ml_EdadMaxDias
End Property
Property Let Restriccion(bValue As Boolean)
   mb_Restriccion = bValue
End Property
Property Get Restriccion() As Boolean
   Restriccion = mb_Restriccion
End Property
Property Let CodigoCPT2004(sValue As String)
   ms_CodigoCPT2004 = sValue
End Property
Property Get CodigoCPT2004() As String
   CodigoCPT2004 = ms_CodigoCPT2004
End Property
Property Let CodigoCPT99(sValue As String)
   ms_CodigoCPT99 = sValue
End Property
Property Get CodigoCPT99() As String
   CodigoCPT99 = ms_CodigoCPT99
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdProcedimiento(lValue As Long)
   ml_IdProcedimiento = lValue
End Property
Property Get IdProcedimiento() As Long
   IdProcedimiento = ml_IdProcedimiento
End Property
Public Property Get IdDiferenciacion() As Integer
    IdDiferenciacion = mi_IdDiferenciacion
End Property
Public Property Let IdDiferenciacion(ByVal iValue As Integer)
    mi_IdDiferenciacion = iValue
End Property
