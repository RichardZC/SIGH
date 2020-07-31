VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Empleados
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Clave As String
Dim ms_USUARIO As String
Dim mda_FechaAlta As Date
Dim mda_FechaIngreso As Date
Dim ms_CodigoPlanilla As String
Dim ms_DNI As String
Dim ml_IdTipoEmpleado As Long
Dim ml_IdCondicionTrabajo As Long
Dim ms_Nombres As String
Dim ms_ApellidoMaterno As String
Dim ms_ApellidoPaterno As String
Dim ml_IdEmpleado As Long
Dim ml_LoginEstado As Long
Dim ms_LoginPC As String
Dim mda_FechaNacimiento As Date
Dim ml_idTipoDestacado As Long
Dim ms_HisCodigoDigitador As String
Dim ml_IdEstablecimientoExterno As Long
Dim mb_ReniecAutorizado As Boolean
Dim ml_IdTipoDocumento As Long
Dim ml_IdSupervisor As Long
Dim ml_EsActivo As Boolean
Dim ml_idTipoSexo As Long
'SCCQ 25/03/2020 Cambio2 Inicio
Dim ml_IdPais As Long
Property Let IdPais(lValue As Long)
   ml_IdPais = lValue
End Property
Property Get IdPais() As Long
   IdPais = ml_IdPais
End Property
'SCCQ 25/03/2020 Cambio2 Fin

Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property

Property Let IdSupervisor(lValue As Long)
   ml_IdSupervisor = lValue
End Property
Property Get IdSupervisor() As Long
   IdSupervisor = ml_IdSupervisor
End Property

Property Let IdTipoDocumento(lValue As Long)
   ml_IdTipoDocumento = lValue
End Property
Property Get IdTipoDocumento() As Long
   IdTipoDocumento = ml_IdTipoDocumento
End Property

Property Let ReniecAutorizado(bValue As Boolean)
   mb_ReniecAutorizado = bValue
End Property
Property Get ReniecAutorizado() As Boolean
   ReniecAutorizado = mb_ReniecAutorizado
End Property


Property Let IdEstablecimientoExterno(lValue As Long)
   ml_IdEstablecimientoExterno = lValue
End Property
Property Get IdEstablecimientoExterno() As Long
   IdEstablecimientoExterno = ml_IdEstablecimientoExterno
End Property
Property Let HisCodigoDigitador(sValue As String)
   ms_HisCodigoDigitador = sValue
End Property
Property Get HisCodigoDigitador() As String
   HisCodigoDigitador = ms_HisCodigoDigitador
End Property


Property Let idTipoDestacado(lValue As Long)
   ml_idTipoDestacado = lValue
End Property
Property Get idTipoDestacado() As Long
   idTipoDestacado = ml_idTipoDestacado
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property

Property Let LoginPc(sValue As String)
   ms_LoginPC = sValue
End Property
Property Get LoginPc() As String
   LoginPc = ms_LoginPC
End Property
Property Let LoginEstado(lValue As Long)
   ml_LoginEstado = lValue
End Property
Property Get LoginEstado() As Long
   LoginEstado = ml_LoginEstado
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Clave(sValue As String)
   ms_Clave = sValue
End Property
Property Get Clave() As String
   Clave = ms_Clave
End Property
Property Let USUARIO(sValue As String)
   ms_USUARIO = sValue
End Property
Property Get USUARIO() As String
   USUARIO = ms_USUARIO
End Property
Property Let FechaAlta(daValue As Date)
   mda_FechaAlta = daValue
End Property
Property Get FechaAlta() As Date
   FechaAlta = mda_FechaAlta
End Property
Property Let FechaIngreso(daValue As Date)
   mda_FechaIngreso = daValue
End Property
Property Get FechaIngreso() As Date
   FechaIngreso = mda_FechaIngreso
End Property
Property Let CodigoPlanilla(sValue As String)
   ms_CodigoPlanilla = sValue
End Property
Property Get CodigoPlanilla() As String
   CodigoPlanilla = ms_CodigoPlanilla
End Property
Property Let DNI(sValue As String)
   ms_DNI = sValue
End Property
Property Get DNI() As String
   DNI = ms_DNI
End Property
Property Let IdTipoEmpleado(lValue As Long)
   ml_IdTipoEmpleado = lValue
End Property
Property Get IdTipoEmpleado() As Long
   IdTipoEmpleado = ml_IdTipoEmpleado
End Property
Property Let IdCondicionTrabajo(lValue As Long)
   ml_IdCondicionTrabajo = lValue
End Property
Property Get IdCondicionTrabajo() As Long
   IdCondicionTrabajo = ml_IdCondicionTrabajo
End Property
Property Let Nombres(sValue As String)
   ms_Nombres = sValue
End Property
Property Get Nombres() As String
   Nombres = ms_Nombres
End Property
Property Let ApellidoMaterno(sValue As String)
   ms_ApellidoMaterno = sValue
End Property
Property Get ApellidoMaterno() As String
   ApellidoMaterno = ms_ApellidoMaterno
End Property
Property Let ApellidoPaterno(sValue As String)
   ms_ApellidoPaterno = sValue
End Property
Property Get ApellidoPaterno() As String
   ApellidoPaterno = ms_ApellidoPaterno
End Property
Property Let IdEmpleado(lValue As Long)
   ml_IdEmpleado = lValue
End Property
Property Get IdEmpleado() As Long
   IdEmpleado = ml_IdEmpleado
End Property
Property Let EsActivo(lValue As Boolean)
   ml_EsActivo = lValue
End Property
Property Get EsActivo() As Boolean
   EsActivo = ml_EsActivo
End Property


