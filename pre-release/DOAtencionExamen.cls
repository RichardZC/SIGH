VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionExamen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesExamenes
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdDetalleProducto As Long
Dim ms_HoraOrden As String
Dim ml_IdServicioOrdena As Long
Dim ml_IdExamen As Long
Dim mda_FechaOrden As Date
Dim ms_OrdenNro As String
Dim ml_IdMedicoOrdena As Long
Dim ml_idCuentaAtencion As Long
Dim ms_HoraResultado As String
Dim mda_FechaResultado As Date
Dim ml_IdAtencionExamen As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdDetalleProducto(lValue As Long)
   ml_IdDetalleProducto = lValue
End Property
Property Get IdDetalleProducto() As Long
   IdDetalleProducto = ml_IdDetalleProducto
End Property
Property Let HoraOrden(sValue As String)
   ms_HoraOrden = sValue
End Property
Property Get HoraOrden() As String
   HoraOrden = ms_HoraOrden
End Property
Property Let IdServicioOrdena(lValue As Long)
   ml_IdServicioOrdena = lValue
End Property
Property Get IdServicioOrdena() As Long
   IdServicioOrdena = ml_IdServicioOrdena
End Property
Property Let IdExamen(lValue As Long)
   ml_IdExamen = lValue
End Property
Property Get IdExamen() As Long
   IdExamen = ml_IdExamen
End Property
Property Let FechaOrden(daValue As Date)
   mda_FechaOrden = daValue
End Property
Property Get FechaOrden() As Date
   FechaOrden = mda_FechaOrden
End Property
Property Let OrdenNro(sValue As String)
   ms_OrdenNro = sValue
End Property
Property Get OrdenNro() As String
   OrdenNro = ms_OrdenNro
End Property
Property Let IdMedicoOrdena(lValue As Long)
   ml_IdMedicoOrdena = lValue
End Property
Property Get IdMedicoOrdena() As Long
   IdMedicoOrdena = ml_IdMedicoOrdena
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let HoraResultado(sValue As String)
   ms_HoraResultado = sValue
End Property
Property Get HoraResultado() As String
   HoraResultado = ms_HoraResultado
End Property
Property Let FechaResultado(daValue As Date)
   mda_FechaResultado = daValue
End Property
Property Get FechaResultado() As Date
   FechaResultado = mda_FechaResultado
End Property
Property Let IdAtencionExamen(lValue As Long)
   ml_IdAtencionExamen = lValue
End Property
Property Get IdAtencionExamen() As Long
   IdAtencionExamen = ml_IdAtencionExamen
End Property


