VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCatalogoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactCatalogoServicios
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdServicioSubGrupo As Long
Dim ml_IdPartida As Long
Dim ml_IdCentroCosto As Long
Dim ml_IdServicioSubSeccion As Long
Dim ml_IdServicioSeccion As Long
Dim ml_IdServicioGrupo As Long
Dim ms_nombre As String
Dim ms_Codigo As String
Dim ml_idProducto As Long
Dim ml_EsCPT As Long
Dim ms_NombreMINSA As String
Dim mb_idEstado As Long
Dim ms_codigoSIS As String
Dim ml_LabResultadoAutomatico As Long
Dim ml_idProveedor As Long
Dim ms_EquipoCodigo As String
Dim ms_codigoSUNAT As String

Property Let codigoSUNAT(sValue As String)
   ms_codigoSUNAT = sValue
End Property
Property Get codigoSUNAT() As String
   codigoSUNAT = ms_codigoSUNAT
End Property


Property Let EquipoCodigo(sValue As String)
   ms_EquipoCodigo = sValue
End Property
Property Get EquipoCodigo() As String
   EquipoCodigo = ms_EquipoCodigo
End Property

Property Let idProveedor(lValue As Long)
   ml_idProveedor = lValue
End Property
Property Get idProveedor() As Long
   idProveedor = ml_idProveedor
End Property

Property Let LabResultadoAutomatico(lValue As Long)
   ml_LabResultadoAutomatico = lValue
End Property
Property Get LabResultadoAutomatico() As Long
   LabResultadoAutomatico = ml_LabResultadoAutomatico
End Property

Property Let codigoSIS(sValue As String)
   ms_codigoSIS = sValue
End Property
Property Get codigoSIS() As String
   codigoSIS = ms_codigoSIS
End Property

Property Let idEstado(lValue As Long)
   mb_idEstado = lValue
End Property
Property Get idEstado() As Long
   idEstado = mb_idEstado
End Property
Property Let NombreMINSA(sValue As String)
   ms_NombreMINSA = sValue
End Property
Property Get NombreMINSA() As String
   NombreMINSA = ms_NombreMINSA
End Property

Property Let EsCpt(lValue As Long)
   ml_EsCPT = lValue
End Property
Property Get EsCpt() As Long
   EsCpt = ml_EsCPT
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdServicioSubGrupo(lValue As Long)
   ml_IdServicioSubGrupo = lValue
End Property
Property Get IdServicioSubGrupo() As Long
   IdServicioSubGrupo = ml_IdServicioSubGrupo
End Property
Property Let IdPartida(lValue As Long)
   ml_IdPartida = lValue
End Property
Property Get IdPartida() As Long
   IdPartida = ml_IdPartida
End Property
Property Let IdCentroCosto(lValue As Long)
   ml_IdCentroCosto = lValue
End Property
Property Get IdCentroCosto() As Long
   IdCentroCosto = ml_IdCentroCosto
End Property
Property Let IdServicioSubSeccion(lValue As Long)
   ml_IdServicioSubSeccion = lValue
End Property
Property Get IdServicioSubSeccion() As Long
   IdServicioSubSeccion = ml_IdServicioSubSeccion
End Property
Property Let IdServicioSeccion(lValue As Long)
   ml_IdServicioSeccion = lValue
End Property
Property Get IdServicioSeccion() As Long
   IdServicioSeccion = ml_IdServicioSeccion
End Property
Property Let IdServicioGrupo(lValue As Long)
   ml_IdServicioGrupo = lValue
End Property
Property Get IdServicioGrupo() As Long
   IdServicioGrupo = ml_IdServicioGrupo
End Property
Property Let nombre(sValue As String)
   ms_nombre = sValue
End Property
Property Get nombre() As String
   nombre = ms_nombre
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let idProducto(lValue As Long)
   ml_idProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property


