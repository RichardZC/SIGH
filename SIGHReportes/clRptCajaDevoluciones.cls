VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clRptCajaDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ml_idUsuario  As Long
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New rptCajaDevoluciones
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

'****************************************
'******************INO*******************
'****************************************


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, _
                   mda_FechaInicio As Date, mda_FechaFin As Date, lnHwnd As Long)
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim ldfechaDevolucion As Date
    Dim ldNroSerie As String, ldNroDocumento As String, ldPaciente As String, ldUsuario As String, ldMotivo As String
    Dim ldTotal As Double, ldTotalAcumulado As Double, ldIdComprobantePago As Integer
    
      Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
    
    ldTotalAcumulado = 0
    
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set mrs_Tmp1 = mo_ReglasCaja.CajaDevolucionesPorFechas(mda_FechaInicio & " 00:01", mda_FechaFin & " 23:59")
    'mrs_Tmp1.Filter = "idListItem=501 or idListItem=502 or idListItem=1330"
    If mrs_Tmp1.RecordCount = 0 Then
         MsgBox "No existe Datos", vbInformation, "Reporte"
    Else
    
          With mrs_Tmp
            .Fields.Append "IdComprobantePago", adInteger, adFldIsNullable
            .Fields.Append "NroSerie", adChar, 4, adFldIsNullable
            .Fields.Append "NroDocumento", adChar, 7, adFldIsNullable
            .Fields.Append "Paciente", adVarChar, 200, adFldIsNullable
            .Fields.Append "Total", adDouble, adFldIsNullable
            .Fields.Append "fechaDevolucion", adDate, 250, adFldIsNullable
            .Fields.Append "Usuario", adVarChar, 20, adFldIsNullable
            .Fields.Append "Motivo", adVarChar, 2000, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
         End With
               
         mrs_Tmp1.MoveFirst
         Do While Not mrs_Tmp1.EOF
         
          ldIdComprobantePago = mrs_Tmp1.Fields!IdComprobantePago
          ldNroSerie = mrs_Tmp1.Fields!nroSerie
          ldNroDocumento = mrs_Tmp1.Fields!NroDocumento
          ldPaciente = mrs_Tmp1.Fields!Paciente
          ldTotal = mrs_Tmp1.Fields!Total
          ldfechaDevolucion = mrs_Tmp1.Fields!fechadevolucion
          ldUsuario = mrs_Tmp1.Fields!Usuario
          ldMotivo = mrs_Tmp1.Fields!motivo
                      
             mrs_Tmp.AddNew
                mrs_Tmp.Fields!IdComprobantePago = ldIdComprobantePago
                mrs_Tmp.Fields!nroSerie = ldNroSerie
                mrs_Tmp.Fields!NroDocumento = ldNroDocumento
                mrs_Tmp.Fields!Paciente = ldPaciente
                mrs_Tmp.Fields!Total = ldTotal
                mrs_Tmp.Fields!fechadevolucion = ldfechaDevolucion
                mrs_Tmp.Fields!Usuario = ldUsuario
                mrs_Tmp.Fields!motivo = ldMotivo
                mrs_Tmp.Update
                
           ldTotalAcumulado = ldTotalAcumulado + ldTotal
           
         mrs_Tmp1.MoveNext
         Loop
         
        ' lcPie = "Nro Movimientos: " & Trim(Str(mrs_Tmp.RecordCount))
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
        Else
        
            With RpCajaDevoluciones
                 .Orientation = rptOrientLandscape
                 .Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                 .Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
'                 Set .Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                 .Sections("pie").Controls("lblPie").Caption = lcPie
                 .Sections("pie").Controls("lblTotalAcumulado").Caption = ldTotalAcumulado
            End With
         
            Set RpCajaDevoluciones.DataSource = mrs_Tmp
'           ' RpCajaDevoluciones.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
'            'RpCajaDevoluciones.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
'           ' Set RpCajaDevoluciones.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
'          '  RpCajaDevoluciones.Sections("pie").Controls("lblPie").Caption = lcPie
            RpCajaDevoluciones.Orientation = rptOrientLandscape
            RpCajaDevoluciones.Show 1
            'debb-27/05/2015
            Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
            mo_reglasComunes.grabaTablaAuditoria ("RpCajaDevoluciones: " & _
                                           Mid(lcSubTitulo, IIf(InStr(lcSubTitulo, "FILTROS: ") > 0, 10, 1)))
            Set mo_reglasComunes = Nothing
        End If
    End If
    Set oConexion = Nothing
End Sub


