VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Proveedores
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idProveedor As Long
Dim ms_RUC As String
Dim ms_RazonSocial As String
Dim ms_Email As String
Dim ms_Direccion As String
Property Let Direccion(sValue As String)
   ms_Direccion = sValue
End Property
Property Get Direccion() As String
   Direccion = ms_Direccion
End Property

Property Let Email(sValue As String)
   ms_Email = sValue
End Property
Property Get Email() As String
   Email = ms_Email
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idProveedor(iValue As Long)
   ml_idProveedor = iValue
End Property
Property Get idProveedor() As Long
   idProveedor = ml_idProveedor
End Property
Property Let RUC(sValue As String)
   ms_RUC = sValue
End Property
Property Get RUC() As String
   RUC = ms_RUC
End Property
Property Let RazonSocial(sValue As String)
   ms_RazonSocial = sValue
End Property
Property Get RazonSocial() As String
   RazonSocial = ms_RazonSocial
End Property

