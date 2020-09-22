VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmPreVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmPreVenta
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idPreventa As Long
Dim ml_idAlmacen As Long
Dim ml_idVendedor As Long
Dim ml_IdPaciente As Long
Dim ml_IdTipoFinanciamiento As Long
Dim mc_Total As Currency
Dim ml_IdDiagnostico As Long
Dim ml_idTipoReceta As Long
Dim ml_IdCuentaAtencion As Long
Dim ml_idPrescriptor As Long
Dim mda_FechaCreacion As Date
Dim ml_idUsuario As Long
Dim mda_FechaModificacion As Date
Dim ml_idUsuarioModifica As Long
Dim ml_idEstadoPreventa As Long
Dim ms_HoraCreacion As String
Dim mda_FechaHoraPrescribe As Date
Dim ms_dni As String
Dim ms_Paciente  As String

Property Let Paciente(sValue As String)
   ms_Paciente = sValue
End Property
Property Get Paciente() As String
   Paciente = ms_Paciente
End Property

Property Let dni(sValue As String)
   ms_dni = sValue
End Property
Property Get dni() As String
   dni = ms_dni
End Property


Property Let FechaHoraPrescribe(daValue As Date)
   mda_FechaHoraPrescribe = daValue
End Property
Property Get FechaHoraPrescribe() As Date
   FechaHoraPrescribe = mda_FechaHoraPrescribe
End Property

Property Let HoraCreacion(sValue As String)
   ms_HoraCreacion = sValue
End Property
Property Get HoraCreacion() As String
   HoraCreacion = ms_HoraCreacion
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let idPreventa(iValue As Long)
   ml_idPreventa = iValue
End Property
Property Get idPreventa() As Long
   idPreventa = ml_idPreventa
End Property
Property Let idAlmacen(iValue As Long)
   ml_idAlmacen = iValue
End Property
Property Get idAlmacen() As Long
   idAlmacen = ml_idAlmacen
End Property
Property Let idVendedor(iValue As Long)
   ml_idVendedor = iValue
End Property
Property Get idVendedor() As Long
   idVendedor = ml_idVendedor
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let TOTAL(cValue As Currency)
   mc_Total = cValue
End Property
Property Get TOTAL() As Currency
   TOTAL = mc_Total
End Property
Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let idTipoReceta(iValue As Long)
   ml_idTipoReceta = iValue
End Property
Property Get idTipoReceta() As Long
   idTipoReceta = ml_idTipoReceta
End Property
Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let idPrescriptor(iValue As Long)
   ml_idPrescriptor = iValue
End Property
Property Get idPrescriptor() As Long
   idPrescriptor = ml_idPrescriptor
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
Property Let FechaModificacion(daValue As Date)
   mda_FechaModificacion = daValue
End Property
Property Get FechaModificacion() As Date
   FechaModificacion = mda_FechaModificacion
End Property
Property Let idUsuarioModifica(iValue As Long)
   ml_idUsuarioModifica = iValue
End Property
Property Get idUsuarioModifica() As Long
   idUsuarioModifica = ml_idUsuarioModifica
End Property
Property Let idEstadoPreventa(iValue As Long)
   ml_idEstadoPreventa = iValue
End Property
Property Get idEstadoPreventa() As Long
   idEstadoPreventa = ml_idEstadoPreventa
End Property

