VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResultadosImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reporte de Tomografía
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************saldos por almacen
'***************Consideraciones:
Option Explicit
Dim oImagResultados As New ImagResultados
Dim ms_Paciente As String
Dim ml_PuntoCarga As Long
Dim oRsResultados As New Recordset
Dim ml_idMovimiento As Long
Dim ml_idProductoCpt As Long
Dim mb_EsResultadoAutomatico As Boolean
Dim ml_Producto As String
Dim mb_SoloEsConsulta As Boolean
Property Let SoloEsConsulta(lValue As Boolean)
    mb_SoloEsConsulta = lValue
End Property

Property Let Producto(lValue As String)
    ml_Producto = lValue
End Property
Property Let EsResultadoAutomatico(lValue As Boolean)
    mb_EsResultadoAutomatico = lValue
    
End Property


Property Let idProductoCpt(lValue As Long)
    ml_idProductoCpt = lValue
End Property
Property Let idMovimiento(lValue As Long)
    ml_idMovimiento = lValue
End Property

Property Set RsResultados(lValue As Recordset)
    Set oRsResultados = lValue
End Property

Property Let Paciente(lValue As String)
    ms_Paciente = lValue
End Property
Property Let PuntoCarga(lValue As Long)
    ml_PuntoCarga = lValue
End Property
Sub MostrarFormulario()

    oImagResultados.SoloEsConsulta = mb_SoloEsConsulta
    oImagResultados.Producto = ml_Producto
    oImagResultados.EsResultadoAutomatico = mb_EsResultadoAutomatico
    oImagResultados.idProductoCpt = ml_idProductoCpt
    oImagResultados.idMovimiento = ml_idMovimiento
    Set oImagResultados.RsResultados = oRsResultados
    oImagResultados.Paciente = ms_Paciente
    oImagResultados.PuntoCarga = ml_PuntoCarga
    oImagResultados.Show 1
End Sub
