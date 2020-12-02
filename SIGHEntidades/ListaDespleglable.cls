VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListaDespleglable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales del control DATACOMBO
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ComboBox As Object
Dim ms_BoundText As String
Dim ms_BoundColumn As String
Dim ms_ListField As String
Dim mrs_RecordSet As ADODB.Recordset
Property Set MiComboBox(oValue As Object)
    Set mo_ComboBox = oValue
End Property
Property Get MiComboBox() As Object
    Set MiComboBox = mo_ComboBox
End Property
Property Get ListField() As String
    ListField = ms_ListField
End Property
Property Let ListField(sValue As String)
    ms_ListField = sValue
End Property
Property Get BoundColumn() As String
    BoundColumn = ms_BoundColumn
End Property
Property Let BoundColumn(sValue As String)
    ms_BoundColumn = sValue
End Property
Property Set RowSource(rsRecordset As ADODB.Recordset)
    If Not rsRecordset Is Nothing Then
        Set mrs_RecordSet = rsRecordset
        CargarComboBoxDesdeRecordset mo_ComboBox, rsRecordset, ms_BoundColumn, ms_ListField
    End If
End Property
Property Get RowSource() As Recordset
    Set RowSource = mrs_RecordSet
End Property
Property Set RowSourceSinClear(rsRecordset As ADODB.Recordset)
    If Not rsRecordset Is Nothing Then
        Set mrs_RecordSet = rsRecordset
        CargarComboBoxDesdeRecordsetSinClear mo_ComboBox, rsRecordset, ms_BoundColumn, ms_ListField
    End If
End Property
Property Let BoundText(sValue As String)
    
    ms_BoundText = sValue
    On Error Resume Next
    UbicarItemDeComboBoxPorId mo_ComboBox, Val(sValue)

End Property
Property Get BoundText() As String
    'BoundText = ms_BoundText
    On Error Resume Next
    If mo_ComboBox.ListIndex <> -1 Then
        BoundText = mo_ComboBox.ItemData(mo_ComboBox.ListIndex)
    Else
        BoundText = ""
    End If
End Property

'------------------------------------------------------------------------------------
'   Función:        UbicarItemDeComboBoxPorId
'   Descripción:    Ubica un Item dentro de unCombo Box, basado en el Id de la propiedad ItemData
'   Parámetros:     oCombo:     Es el Combo Box que se va a llenar
'                   iId:        Es el Id que se desea ubicar.
'------------------------------------------------------------------------------------

Sub UbicarItemDeComboBoxPorId(oCombo As Object, lId As Long)
Dim I As Integer

    For I = 0 To oCombo.ListCount - 1
        If oCombo.ItemData(I) = lId Then
            oCombo.ListIndex = I
            Exit Sub
        End If
    Next I
    oCombo.ListIndex = -1
    
End Sub
'------------------------------------------------------------------------------------
'   Función:        ObtenerItemDataDeComboxBox
'   Descripción:    Obtiene el valor de la propiedad ItemData de un ComboBox
'                   basado en el valor seleccionado.
'   Parámetros:     oCombo:     Es el Combo Box que se va a llenar
'------------------------------------------------------------------------------------
Function ObtenerItemDataDeComboxBox(oCombo As Object) As Long

    If oCombo.ListIndex = -1 Then
        ObtenerItemDataDeComboxBox = -1
    Else
        ObtenerItemDataDeComboxBox = oCombo.ItemData(oCombo.ListIndex)
    End If

End Function
'------------------------------------------------------------------------------------
'   Función:        UbicarItemDeComboBoxPorTexto
'   Descripción:    Ubica un Item dentro de unCombo Box, basado en el texto del combo
'   Parámetros:     oDataCombo:     Es el Combo Box que se va a llenar
'                   oParametros:        Es el objeto que almacena los parametros del combo
'------------------------------------------------------------------------------------
Sub UbicarItemDeComboBoxPorTexto(oDataCombo As Object)
Dim oCadena As New Cadena

    If InStr(oDataCombo.Text, "=") > 0 Then
        oDataCombo.BoundText = Trim(oCadena.ExtraerCadena(oDataCombo.Text, 1, "="))
    Else
        oDataCombo.BoundText = Trim(oDataCombo.Text)
    End If

    Set oCadena = Nothing

End Sub

Sub CargarComboBoxDesdeRecordset(oComboBox As Object, rsDatos As ADODB.Recordset, sCodigo As String, sDescripcion As String)
        oComboBox.Clear
        On Error Resume Next
        rsDatos.MoveFirst
        Do While Not rsDatos.EOF
            If Not IsNull(rsDatos(sDescripcion)) Then
                oComboBox.AddItem rsDatos(sDescripcion)
                oComboBox.ItemData(oComboBox.NewIndex) = rsDatos(sCodigo)
            End If
            rsDatos.MoveNext
        Loop
End Sub

Sub CargarComboBoxDesdeRecordsetSinClear(oComboBox As Object, rsDatos As ADODB.Recordset, sCodigo As String, sDescripcion As String)
        Do While Not rsDatos.EOF
        oComboBox.AddItem rsDatos(sDescripcion)
        oComboBox.ItemData(oComboBox.NewIndex) = rsDatos(sCodigo)
        rsDatos.MoveNext
        Loop
End Sub

'mgaray201411f
Property Set RowSourceTextCompuesto(rsRecordset As ADODB.Recordset)
    If Not rsRecordset Is Nothing Then
        Set mrs_RecordSet = rsRecordset
        CargarComboBoxDesdeRecordsetTextCompuesto mo_ComboBox, rsRecordset, ms_BoundColumn, ms_ListField
    End If
End Property
'permite concatenar varios campos en lo que se mmuestra en la lista de la lista desplegable
'|:fieldName|(|:fieldName|) los valores que desea concatenar deben ir separados por | y si se quiere concatener el valor de un campo de recordset se debe iniciar con :
Private Sub CargarComboBoxDesdeRecordsetTextCompuesto(oComboBox As Object, rsDatos As ADODB.Recordset, sCodigo As String, sDescripcion As String)
        oComboBox.Clear
        On Error Resume Next
        rsDatos.MoveFirst
        
        Dim arrayListaCampos As Variant
        
        
        arrayListaCampos = Split(sDescripcion, "|")
        If UBound(arrayListaCampos) = 0 Then
            CargarComboBoxDesdeRecordset mo_ComboBox, rsDatos, ms_BoundColumn, ms_ListField
            Exit Sub
        End If
        
        Do While Not rsDatos.EOF
'            If Not IsNull(rsDatos(sDescripcion)) Then
                oComboBox.AddItem DevuelveValorTextoCompuesto(rsDatos, arrayListaCampos) 'rsDatos(sDescripcion)
                oComboBox.ItemData(oComboBox.NewIndex) = rsDatos(sCodigo)
'            End If
            rsDatos.MoveNext
        Loop
End Sub
Private Function DevuelveValorTextoCompuesto(rsDatos As ADODB.Recordset, arrayListaCampos As Variant) As String
    Dim I As Integer
    Dim sCadena As String
    Dim sFieldName As String
    
    sCadena = ""
    For I = 0 To UBound(arrayListaCampos)
        
        If Left(arrayListaCampos(I), 1) = ":" Then
            sFieldName = Mid(arrayListaCampos(I), 2)
            sCadena = sCadena & Trim(IIf(IsNull(rsDatos(sFieldName)), "", rsDatos(sFieldName)))
        Else
            sCadena = sCadena & arrayListaCampos(I)
        End If
    Next I
    
    DevuelveValorTextoCompuesto = sCadena
End Function
