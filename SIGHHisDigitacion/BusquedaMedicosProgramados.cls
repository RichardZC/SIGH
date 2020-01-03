VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaMedicosHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Médicos HIS
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim oFormulario As New MedicosHISBusqueda
Dim ml_IdRegistroSeleccionado As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdEstablecimiento As Long
Dim ml_IdServicio As Long
Dim ml_IdTurno As Integer
Dim ml_IdMedico As Long
Dim ml_NombreMedico  As String
Dim ml_Anio  As String
Dim ml_Mes  As String

Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let Anio(lValue As String)
    ml_Anio = lValue
End Property
Property Get Anio() As String
    Anio = ml_Anio
End Property
Property Let Mes(lValue As String)
    ml_Mes = lValue
End Property
Property Get Mes() As String
    Mes = ml_Mes
End Property
Property Let IdServicio(lValue As Long)
    ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
    IdServicio = ml_IdServicio
End Property
Property Let IdTurno(lValue As Integer)
    ml_IdTurno = lValue
End Property
Property Get IdTurno() As Integer
    IdTurno = ml_IdTurno
End Property
Property Let IdMedico(lValue As Long)
    ml_IdMedico = lValue
End Property
Property Get IdMedico() As Long
    IdMedico = ml_IdMedico
End Property
Property Let NombMedico(lValue As String)
    ml_NombreMedico = lValue
End Property
Property Get NombMedico() As String
    NombMedico = ml_NombreMedico
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property

Sub MostrarFormulario()
    oFormulario.IdEstablecimiento = ml_IdEstablecimiento
'    oFormulario.idIdServicio = ldFechaProgramada
'    oFormulario.HoraProgramada = lcHoraProgramada
'    oFormulario.idTipoServicio = lnIdTipoServicio
    oFormulario.NombreMedico = ml_NombreMedico
    oFormulario.Anio = ml_Anio
    oFormulario.Mes = ml_Mes
    oFormulario.Show 1
    ml_IdServicio = oFormulario.IdServicio
    ml_IdTurno = oFormulario.IdTurno
    ml_IdMedico = oFormulario.IdMedico
    ml_NombreMedico = oFormulario.NombreMedico
    mi_BotonPresionado = oFormulario.BotonPresionado
    Set oFormulario = Nothing
End Sub
