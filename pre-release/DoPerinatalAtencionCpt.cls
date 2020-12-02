VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoPerinatalAtencionCpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla PerinatalAtencionProcedimientos
'        Programado por: Garay M
'        Fecha: Setiembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPerinatalAtencion As Long
Dim ml_idModulo As Long
Dim ml_idLista As Long
Dim ml_IdProducto As Long
Dim mb_CptEsAutomatico As Boolean
Dim ms_CodigoHIS As String
Dim ml_IdAtencion As Long
Dim ml_IdOrden As Long
'mgaray201411a
Dim ms_LabConfHIS As String
Dim ml_ItemCpt As Long

Property Let IdAtencion(iValue As Long)
   ml_IdAtencion = iValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idPerinatalAtencion(iValue As Long)
   ml_idPerinatalAtencion = iValue
End Property
Property Get idPerinatalAtencion() As Long
   idPerinatalAtencion = ml_idPerinatalAtencion
End Property
Property Let idModulo(iValue As Long)
   ml_idModulo = iValue
End Property
Property Get idModulo() As Long
   idModulo = ml_idModulo
End Property
Property Let idLista(iValue As Long)
   ml_idLista = iValue
End Property
Property Get idLista() As Long
   idLista = ml_idLista
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let CptEsAutomatico(bValue As Boolean)
   mb_CptEsAutomatico = bValue
End Property
Property Get CptEsAutomatico() As Boolean
   CptEsAutomatico = mb_CptEsAutomatico
End Property
Property Let CodigoHIS(sValue As String)
   ms_CodigoHIS = sValue
End Property
Property Get CodigoHIS() As String
   CodigoHIS = ms_CodigoHIS
End Property
Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property
Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property
'mgaray201411a
Property Let labConfHIS(sValue As String)
   ms_LabConfHIS = sValue
End Property
Property Get labConfHIS() As String
   labConfHIS = ms_LabConfHIS
End Property
Property Let ItemCpt(iValue As Long)
   ml_ItemCpt = iValue
End Property
Property Get ItemCpt() As Long
   ItemCpt = ml_ItemCpt
End Property
