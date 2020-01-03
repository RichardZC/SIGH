VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridInfragistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales del control Infragistics
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Public Sub ConfigurarFilasBiColores(grdEmpleados As Object, sBicolor As String)

    On Error Resume Next
    If sBicolor = 99 Then
        grdEmpleados.Bands(0).Override.RowAlternateAppearance.BackColor = vbGreen
        'grdEmpleados.Bands(0).Override.RowAppearance.BackColor = vbGreen
        grdEmpleados.Font = "Arial Narrow"
        grdEmpleados.Appeareance.Font.Size = 8
        grdEmpleados.Bands(1).Override.RowAlternateAppearance.BackColor = vbGreen
        grdEmpleados.Override.BorderStyleCell = ssBorderStyleInset
        grdEmpleados.Override.BorderStyleRow = ssBorderStyleRaisedSoft
       Exit Sub
    End If
    With grdEmpleados
        .Appeareance.Font = "Tahoma"
        .Appeareance.Font.Size = "12"
        .Appeareance.BorderColor = &H808080
        
        .Override.BorderStyleCell = 2
        .Override.BorderStyleHeader = 2
        .Override.BorderStyleRow = 2
        
        .Override.CellAppearance.BorderColor = &H808080
        .Override.RowAppearance.BorderColor = &H808080
        
        .Override.HeaderAppearance.BackColor = &HEEB197   '&H81563D
        .Override.HeaderAppearance.BorderColor = &HFFFFFF               '&H808080
        .Override.HeaderAppearance.ForeColor = &HFFFFFF
        .Override.RowAlternateAppearance.BackColor = &HB8EEFE
    End With

    If sBicolor = 1 Then
        grdEmpleados.Bands(0).Override.RowAlternateAppearance.BackColor = &HFFFFFF
        grdEmpleados.Bands(0).Override.RowAppearance.BackColor = &HFDF0E6
        
        grdEmpleados.Bands(1).Override.RowAlternateAppearance.BackColor = &HFFFFFF
        grdEmpleados.Bands(1).Override.RowAppearance.BackColor = &HDAFDFE
    ElseIf sBicolor = 99 Then
        grdEmpleados.Bands(0).Override.RowAlternateAppearance.BackColor = vbGreen
        grdEmpleados.Bands(0).Override.RowAppearance.BackColor = vbGreen
        grdEmpleados.Bands(1).Override.RowAlternateAppearance.BackColor = vbGreen
        grdEmpleados.Bands(1).Override.RowAppearance.BackColor = vbGreen
    
    Else
        grdEmpleados.Bands(0).Override.RowAlternateAppearance.BackColor = &HFFFFFF
        grdEmpleados.Bands(0).Override.RowAppearance.BackColor = &HFFFFFF
    End If
        
End Sub
Public Sub ConfigurarFilasAlphaBlending(grdEmpleados As Object, sRutaImagen As String)
        
        grdEmpleados.Appearance.PictureBackground = LoadPicture(sRutaImagen)
        grdEmpleados.Appearance.PictureBackgroundStyle = ssPictureBackgroundStyleTiled
        
        Set_Appearance_UseAlpha grdEmpleados.Override.HeaderAppearance, 192, 1
        Set_Appearance_UseAlpha grdEmpleados.Override.RowSelectorAppearance, 192, 1
        Set_Appearance_UseAlpha grdEmpleados.Override.RowAppearance, 128, 1
        Set_Appearance_UseAlpha grdEmpleados.Override.CellAppearance, 129, 1
        grdEmpleados.Override.CellAppearance.ForegroundAlpha = 2

End Sub
Private Sub Set_Appearance_UseAlpha(App As Object, AlphaLevel As Long, Use As Integer)
    With App
        .AlphaLevel = AlphaLevel
        .BackColorAlpha = Use
        .BorderAlpha = Use
        .ForegroundAlpha = Use
        .PictureAlpha = Use
        .PictureBackgroundAlpha = Use
    End With
    
End Sub
'mgaray
Public Function ocultarColumnas(ByVal Layout As UltraGrid.SSLayout, ByVal indexBand, _
                    ParamArray ParamArrayArgs() As Variant) As Boolean
    Dim I As Integer
    
    With Layout.Bands(indexBand)
        For I = 0 To UBound(ParamArrayArgs)
            .Columns(ParamArrayArgs(I)).Hidden = True
        Next I
    End With
End Function

Public Function modificarActivationColumnas(ByVal Layout As UltraGrid.SSLayout, ByVal indexBand, _
                    optionActivation As Constants_Activation, ParamArray ParamArrayArgs() As Variant) As Boolean
    Dim I As Integer
    
    With Layout.Bands(indexBand)
        For I = 0 To UBound(ParamArrayArgs)
            .Columns(ParamArrayArgs(I)).Activation = optionActivation
        Next I
    End With
End Function

Public Function modificarAlineacionHColumnas(ByVal Layout As UltraGrid.SSLayout, ByVal indexBand, _
                    optionAlign As Constants_Align, ParamArray ParamArrayArgs() As Variant) As Boolean
    Dim I As Integer
    
    With Layout.Bands(indexBand)
        For I = 0 To UBound(ParamArrayArgs)
            .Columns(ParamArrayArgs(I)).CellAppearance.TextAlign = optionAlign
        Next I
    End With
End Function



