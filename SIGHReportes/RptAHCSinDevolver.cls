VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHCSinDevolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SCCQ 28/08/2020 Cambio27 Inicio
Dim ml_mostrarReporte As Boolean

Property Let mostrarReporte(lValue As Boolean)
    ml_mostrarReporte = lValue
End Property
Sub EjecutaFormulario()
    Dim oFormulario As New AHCSinDevolver
    oFormulario.mostrarReporte = ml_mostrarReporte
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub
'SCCQ 28/08/2020 Cambio27 Fin
