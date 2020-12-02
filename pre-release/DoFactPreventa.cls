VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFactPreventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactPreventa
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idFactPreventa As Long
Dim ml_IdServicio As Long
Dim mc_Total As Currency
Dim mda_FechaCreacion As Date
Dim ml_idUsuario As Long
Dim ml_idEstadoPreventa As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdAtencion As Long
Dim ml_idOrden As Long
Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idFactPreventa(iValue As Long)
   ml_idFactPreventa = iValue
End Property
Property Get idFactPreventa() As Long
   idFactPreventa = ml_idFactPreventa
End Property
Property Let IdServicio(iValue As Long)
   ml_IdServicio = iValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property
Property Let Total(cValue As Currency)
   mc_Total = cValue
End Property
Property Get Total() As Currency
   Total = mc_Total
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idEstadoPreventa(iValue As Long)
   ml_idEstadoPreventa = iValue
End Property
Property Get idEstadoPreventa() As Long
   idEstadoPreventa = ml_idEstadoPreventa
End Property

