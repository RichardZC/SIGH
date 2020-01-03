VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordsetMananger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Sub EjecutarDesconectado(rsRecordset As ADODB.Recordset, oConexion As ADODB.Connection, sSource As String)
'
'    With rsRecordset
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .Source = sSource
'        .ActiveConnection = oConexion
'        .Open
'        .ActiveConnection = Nothing
'    End With
'
'End Sub
