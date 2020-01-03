VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmInventario
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_idInventario As Long
Dim ml_idAlmacen As Long
Dim ms_numeroInventario As String
Dim mda_FechaCierre As Date
Dim mda_FechaCreacion As Date
Dim mda_FechaModificacion As Date
Dim ml_idEstadoInventario As Long
Dim ml_idUsuario As Long
Dim ml_idTipoInventario As Long

Property Let idTipoInventario(iValue As Long)
   ml_idTipoInventario = iValue
End Property
Property Get idTipoInventario() As Long
   idTipoInventario = ml_idTipoInventario
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let idInventario(iValue As Long)
   ml_idInventario = iValue
End Property
Property Get idInventario() As Long
   idInventario = ml_idInventario
End Property
Property Let idAlmacen(iValue As Long)
   ml_idAlmacen = iValue
End Property
Property Get idAlmacen() As Long
   idAlmacen = ml_idAlmacen
End Property
Property Let numeroInventario(sValue As String)
   ms_numeroInventario = sValue
End Property
Property Get numeroInventario() As String
   numeroInventario = ms_numeroInventario
End Property
Property Let FechaCierre(daValue As Date)
   mda_FechaCierre = daValue
End Property
Property Get FechaCierre() As Date
   FechaCierre = mda_FechaCierre
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let FechaModificacion(daValue As Date)
   mda_FechaModificacion = daValue
End Property
Property Get FechaModificacion() As Date
   FechaModificacion = mda_FechaModificacion
End Property
Property Let idEstadoInventario(iValue As Long)
   ml_idEstadoInventario = iValue
End Property
Property Get idEstadoInventario() As Long
   idEstadoInventario = ml_idEstadoInventario
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property


