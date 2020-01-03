VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleProgHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Detalle de Programación HIS
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim ml_IdUsuario As Long
Dim ms_IdMedico As String
Dim mi_IdEspecialidad As Integer
Dim ms_FechaInicial As String
Dim ms_FechaFinal As String
Dim ms_DescripcionEspecialidad As String
Dim ms_IdHisProgMedEstMR As Long
Dim mi_Opcion As sghOpciones
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_FormularioProgramacion As New frmDetalleProgramacion

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let FechaInicial(sValue As String)
   ms_FechaInicial = sValue
End Property

Property Let FechaFinal(sValue As String)
   ms_FechaFinal = sValue
End Property

Property Let IdMedico(sValue As String)
   ms_IdMedico = sValue
End Property

Property Let IdEspecialidad(iValue As Integer)
   mi_IdEspecialidad = iValue
End Property

Property Let DescripcionEspecialidad(sValue As String)
   ms_DescripcionEspecialidad = sValue
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property

Property Let IdHisProgMedEstMR(iValue As Long)
   ms_IdHisProgMedEstMR = iValue
End Property

Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Sub MostrarFormulario()
    mo_FormularioProgramacion.Opcion = mi_Opcion
    mo_FormularioProgramacion.IdUsuario = ml_IdUsuario
    mo_FormularioProgramacion.IdMedico = ms_IdMedico
    mo_FormularioProgramacion.IdEspecialidad = mi_IdEspecialidad
    mo_FormularioProgramacion.FechaInicial = ms_FechaInicial
    mo_FormularioProgramacion.FechaFinal = ms_FechaFinal
    mo_FormularioProgramacion.DescripcionEspecialidad = ms_DescripcionEspecialidad
    mo_FormularioProgramacion.IdHisProgMedEstMR = ms_IdHisProgMedEstMR
    mo_FormularioProgramacion.Show 1
    mi_BotonPresionado = mo_FormularioProgramacion.BotonPresionado
End Sub
