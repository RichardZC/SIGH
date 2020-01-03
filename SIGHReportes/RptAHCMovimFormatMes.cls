VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHCMovimFormatMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Movimiento de Entrada y Salida mensual de Formatos de Historia
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Sub EjecutaFormulario()
    MsgBox "... reporte en desarrollo..."
    Exit Sub

    Dim oFormulario As New AHCMovimFormatMes
    oFormulario.Show 1
End Sub

Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, ml_idTipoServicio As Long, mda_FechaInicio As Date, mda_FechaFin As Date, lnHwnd As Long)
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim mrs_Tmp3 As New Recordset
    Dim mrs_Cab As New Recordset
    Dim mrs_Det As New Recordset
    Dim mrs_Shape As New Recordset
    Dim ldFecha As Date, ldFechaSolo As Date
    Dim lcHoraSolo As String, lcUsuario As String, lcNusuario As String, lcAccion As String
    Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim lnIdListItem As Long, lnNroCuenta As Long, ldFechaIngreso As Date
    Dim lcHoraIngreso As String, lcPaciente As String, lcSexo As String
    Dim lcMotivo As String, lcEstancia As String, ldFechaEgreso As Date
    Dim lcHoraEgreso As String, lcDestino As String, lcRecibe As String
    Dim lcMedicoIE As String, lnImporte As Double, lcDescripcion As String
    Dim lcNombreTablaCab As String, lcNombreTablaDet As String
    Dim lnIdServicio As Long
    Dim lbContinua As Boolean
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
    Dim lnD1 As Integer, lnD2 As Integer, lnD3 As Integer, lnD4 As Integer, lnD5 As Integer
    Dim lnD6 As Integer, lnD7 As Integer, lnD8 As Integer, lnD9 As Integer, lnD10 As Integer
    Dim lnD11 As Integer, lnD12 As Integer, lnD13 As Integer, lnD14 As Integer, lnD15 As Integer
    Dim lnD16 As Integer, lnD17 As Integer, lnD18 As Integer, lnD19 As Integer, lnD20 As Integer
    Dim lnD21 As Integer, lnD22 As Integer, lnD23 As Integer, lnD24 As Integer, lnD25 As Integer
    Dim lnD26 As Integer, lnD27 As Integer, lnD28 As Integer, lnD29 As Integer, lnD30 As Integer
    Dim lnD31 As Integer
    Dim lnDemoraEntregaHC As Integer, lcDemora As String
    Set mrs_Tmp1 = mo_ReglasArchivoClinico.MovimientosFormatoHCMovimientosRetornoHCPorFechasYtipoServicio(mda_FechaInicio, mda_FechaFin, 3)
    If mrs_Tmp1.RecordCount = 0 Then
         MsgBox "No existe Datos", vbInformation, "Reporte"
    Else
         If lbEnExcel = True Then
            'Crear Tabla Temporal CabeceraDetalle para EXCEL --->>>>>>cuando se usa EXCEL
            With mrs_Tmp
               .Fields.Append "idServicio", adInteger
               .Fields.Append "ServicioHosp", adVarChar, 100, adFldIsNullable
               .Fields.Append "DevolucionHC", adVarChar, 30, adFldIsNullable
               .Fields.Append "dia1", adInteger
               .Fields.Append "dia2", adInteger
               .Fields.Append "dia3", adInteger
               .Fields.Append "dia4", adInteger
               .Fields.Append "dia5", adInteger
               .Fields.Append "dia6", adInteger
               .Fields.Append "dia7", adInteger
               .Fields.Append "dia8", adInteger
               .Fields.Append "dia9", adInteger
               .Fields.Append "dia10", adInteger
               .Fields.Append "dia11", adInteger
               .Fields.Append "dia12", adInteger
               .Fields.Append "dia13", adInteger
               .Fields.Append "dia14", adInteger
               .Fields.Append "dia15", adInteger
               .Fields.Append "dia16", adInteger
               .Fields.Append "dia17", adInteger
               .Fields.Append "dia18", adInteger
               .Fields.Append "dia19", adInteger
               .Fields.Append "dia20", adInteger
               .Fields.Append "dia21", adInteger
               .Fields.Append "dia22", adInteger
               .Fields.Append "dia23", adInteger
               .Fields.Append "dia24", adInteger
               .Fields.Append "dia25", adInteger
               .Fields.Append "dia26", adInteger
               .Fields.Append "dia27", adInteger
               .Fields.Append "dia28", adInteger
               .Fields.Append "dia29", adInteger
               .Fields.Append "dia30", adInteger
               .Fields.Append "dia31", adInteger
               .Fields.Append "Total", adInteger
               .LockType = adLockOptimistic
               .Open
            End With
         Else
            'Tablas Cabecera y Detalle para SHAPE--->>>>>> cuando se usa DataReport
            mo_AdminReportes.PreparaTablasDeReportes mrs_Cab, mrs_Det
         End If
         '
         lnImporte = 0
         mrs_Tmp1.MoveFirst
         Do While Not mrs_Tmp1.EOF
            lnIdServicio = mrs_Tmp1.Fields!IdServicioOrigen
            lcTexto1 = mrs_Tmp1.Fields!ServicioOrigen
            If lbEnExcel = False Then
                mrs_Cab.AddNew
                mrs_Cab.Fields!nroCuenta = lnIdServicio
                mrs_Cab.Fields!idServicio = lnIdServicio
                mrs_Cab.Fields!ServicioHosp = lcTexto1
                mrs_Cab.Update
            End If
            Do While Not mrs_Tmp1.EOF And lnIdServicio = mrs_Tmp1.Fields!IdServicioOrigen
                If Not IsNull(mrs_Tmp1.Fields!FechaEgreso) Then
                    lnD1 = 0: lnD2 = 0: lnD3 = 0: lnD4 = 0: lnD5 = 0
                    lnD6 = 0: lnD7 = 0: lnD8 = 0: lnD9 = 0: lnD10 = 0
                    lnD11 = 0: lnD12 = 0: lnD13 = 0: lnD14 = 0: lnD15 = 0
                    lnD16 = 0: lnD17 = 0: lnD18 = 0: lnD19 = 0: lnD20 = 0
                    lnD21 = 0: lnD22 = 0: lnD23 = 0: lnD24 = 0: lnD25 = 0
                    lnD26 = 0: lnD27 = 0: lnD28 = 0: lnD29 = 0: lnD30 = 0
                    lnD31 = 0
                    Select Case Day(mrs_Tmp1.Fields!FechaMovimiento)
                    Case 1
                         lnD1 = lnD1 + 1
                    Case 2
                         lnD2 = lnD2 + 1
                    Case 3
                         lnD3 = lnD3 + 1
                    Case 4
                         lnD4 = lnD4 + 1
                    Case 5
                         lnD5 = lnD5 + 1
                    Case 6
                         lnD6 = lnD6 + 1
                    Case 7
                         lnD7 = lnD7 + 1
                    Case 8
                         lnD8 = lnD8 + 1
                    Case 9
                         lnD9 = lnD9 + 1
                    Case 10
                         lnD10 = lnD10 + 1
                    Case 11
                         lnD11 = lnD11 + 1
                    Case 12
                         lnD12 = lnD12 + 1
                    Case 13
                         lnD13 = lnD13 + 1
                    Case 14
                         lnD14 = lnD14 + 1
                    Case 15
                         lnD15 = lnD15 + 1
                    Case 16
                         lnD16 = lnD16 + 1
                    Case 17
                         lnD17 = lnD17 + 1
                    Case 18
                         lnD18 = lnD18 + 1
                    Case 19
                         lnD19 = lnD19 + 1
                    Case 20
                         lnD20 = lnD20 + 1
                    Case 21
                         lnD21 = lnD21 + 1
                    Case 22
                         lnD22 = lnD22 + 1
                    Case 23
                         lnD23 = lnD23 + 1
                    Case 24
                         lnD24 = lnD24 + 1
                    Case 25
                         lnD25 = lnD25 + 1
                    Case 26
                         lnD26 = lnD26 + 1
                    Case 27
                         lnD27 = lnD27 + 1
                    Case 28
                         lnD28 = lnD28 + 1
                    Case 29
                         lnD29 = lnD29 + 1
                    Case 30
                         lnD30 = lnD30 + 1
                    Case 31
                         lnD31 = lnD31 + 1
                    End Select
                    '
                    lnDemoraEntregaHC = DateDiff("h", mrs_Tmp1.Fields!FechaEgreso & " " & mrs_Tmp1.Fields!horaEgreso, mrs_Tmp1.Fields!FechaMovimiento)
                    If lnDemoraEntregaHC <= 24 Then
                       lcDemora = "<=24hr"
                    ElseIf lnDemoraEntregaHC <= 48 Then
                       lcDemora = ">24 y <=48hr"
                    ElseIf lnDemoraEntregaHC <= 72 Then
                       lcDemora = ">48 y <=72hr"
                    Else
                       lcDemora = ">72hr"
                    End If
                    '
                    If lbEnExcel = False Then
                        lbContinua = True
                        If lnImporte > 0 Then
                           mrs_Det.MoveFirst
                           Do While Not mrs_Det.EOF
                              If mrs_Det.Fields!nroCuenta = lnIdServicio And Trim(mrs_Det.Fields!DevolucionHC) = Trim(lcDemora) Then
                                 lbContinua = False
                                 Exit Do
                              End If
                              mrs_Det.MoveNext
                           Loop
                        End If
                        lnImporte = lnImporte + 1
                        If lbContinua = True Then
                            mrs_Det.AddNew
                            mrs_Det.Fields!nroCuenta = lnIdServicio
                            mrs_Det.Fields!DevolucionHC = lcDemora
                        End If
                        mrs_Det.Fields!dia1 = mrs_Det.Fields!dia1 + lnD1
                        mrs_Det.Fields!dia2 = mrs_Det.Fields!dia2 + lnD2
                        mrs_Det.Fields!dia3 = mrs_Det.Fields!dia3 + lnD3
                        mrs_Det.Fields!dia4 = mrs_Det.Fields!dia4 + lnD4
                        mrs_Det.Fields!dia5 = mrs_Det.Fields!dia5 + lnD5
                        mrs_Det.Fields!dia6 = mrs_Det.Fields!dia6 + lnD6
                        mrs_Det.Fields!dia7 = mrs_Det.Fields!dia7 + lnD7
                        mrs_Det.Fields!dia8 = mrs_Det.Fields!dia8 + lnD8
                        mrs_Det.Fields!dia9 = mrs_Det.Fields!dia9 + lnD9
                        mrs_Det.Fields!dia10 = mrs_Det.Fields!dia10 + lnD10
                        mrs_Det.Fields!dia11 = mrs_Det.Fields!dia11 + lnD11
                        mrs_Det.Fields!dia12 = mrs_Det.Fields!dia12 + lnD12
                        mrs_Det.Fields!dia13 = mrs_Det.Fields!dia13 + lnD13
                        mrs_Det.Fields!dia14 = mrs_Det.Fields!dia14 + lnD14
                        mrs_Det.Fields!dia15 = mrs_Det.Fields!dia15 + lnD15
                        mrs_Det.Fields!dia16 = mrs_Det.Fields!dia16 + lnD16
                        mrs_Det.Fields!dia17 = mrs_Det.Fields!dia17 + lnD17
                        mrs_Det.Fields!dia18 = mrs_Det.Fields!dia18 + lnD18
                        mrs_Det.Fields!dia19 = mrs_Det.Fields!dia19 + lnD19
                        mrs_Det.Fields!dia20 = mrs_Det.Fields!dia20 + lnD20
                        mrs_Det.Fields!dia21 = mrs_Det.Fields!dia21 + lnD21
                        mrs_Det.Fields!dia22 = mrs_Det.Fields!dia22 + lnD22
                        mrs_Det.Fields!dia23 = mrs_Det.Fields!dia23 + lnD23
                        mrs_Det.Fields!dia24 = mrs_Det.Fields!dia24 + lnD24
                        mrs_Det.Fields!dia25 = mrs_Det.Fields!dia25 + lnD25
                        mrs_Det.Fields!dia26 = mrs_Det.Fields!dia26 + lnD26
                        mrs_Det.Fields!dia27 = mrs_Det.Fields!dia27 + lnD27
                        mrs_Det.Fields!dia28 = mrs_Det.Fields!dia28 + lnD28
                        mrs_Det.Fields!dia29 = mrs_Det.Fields!dia29 + lnD29
                        mrs_Det.Fields!dia30 = mrs_Det.Fields!dia30 + lnD30
                        mrs_Det.Fields!dia31 = mrs_Det.Fields!dia31 + lnD31
                        mrs_Det.Fields!Total = mrs_Det.Fields!Total + 1
                        mrs_Det.Update
                    Else
                        lbContinua = True
                        If mrs_Tmp.RecordCount > 0 Then
                           mrs_Tmp.MoveFirst
                           Do While Not mrs_Tmp.EOF
                              If mrs_Tmp.Fields!idServicio = lnIdServicio And Trim(mrs_Tmp.Fields!DevolucionHC) = Trim(lcDemora) Then
                                 lbContinua = False
                                 Exit Do
                              End If
                              mrs_Tmp.MoveNext
                           Loop
                        End If
                        If lbContinua = True Then
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!idServicio = lnIdServicio
                            mrs_Tmp.Fields!ServicioHosp = lcTexto1
                            mrs_Tmp.Fields!DevolucionHC = lcDemora
                        End If
                        mrs_Tmp.Fields!dia1 = mrs_Tmp.Fields!dia1 + lnD1
                        mrs_Tmp.Fields!dia2 = mrs_Tmp.Fields!dia2 + lnD2
                        mrs_Tmp.Fields!dia3 = mrs_Tmp.Fields!dia3 + lnD3
                        mrs_Tmp.Fields!dia4 = mrs_Tmp.Fields!dia4 + lnD4
                        mrs_Tmp.Fields!dia5 = mrs_Tmp.Fields!dia5 + lnD5
                        mrs_Tmp.Fields!dia6 = mrs_Tmp.Fields!dia6 + lnD6
                        mrs_Tmp.Fields!dia7 = mrs_Tmp.Fields!dia7 + lnD7
                        mrs_Tmp.Fields!dia8 = mrs_Tmp.Fields!dia8 + lnD8
                        mrs_Tmp.Fields!dia9 = mrs_Tmp.Fields!dia9 + lnD9
                        mrs_Tmp.Fields!dia10 = mrs_Tmp.Fields!dia10 + lnD10
                        mrs_Tmp.Fields!dia11 = mrs_Tmp.Fields!dia11 + lnD11
                        mrs_Tmp.Fields!dia12 = mrs_Tmp.Fields!dia12 + lnD12
                        mrs_Tmp.Fields!dia13 = mrs_Tmp.Fields!dia13 + lnD13
                        mrs_Tmp.Fields!dia14 = mrs_Tmp.Fields!dia14 + lnD14
                        mrs_Tmp.Fields!dia15 = mrs_Tmp.Fields!dia15 + lnD15
                        mrs_Tmp.Fields!dia16 = mrs_Tmp.Fields!dia16 + lnD16
                        mrs_Tmp.Fields!dia17 = mrs_Tmp.Fields!dia17 + lnD17
                        mrs_Tmp.Fields!dia18 = mrs_Tmp.Fields!dia18 + lnD18
                        mrs_Tmp.Fields!dia19 = mrs_Tmp.Fields!dia19 + lnD19
                        mrs_Tmp.Fields!dia20 = mrs_Tmp.Fields!dia20 + lnD20
                        mrs_Tmp.Fields!dia21 = mrs_Tmp.Fields!dia21 + lnD21
                        mrs_Tmp.Fields!dia22 = mrs_Tmp.Fields!dia22 + lnD22
                        mrs_Tmp.Fields!dia23 = mrs_Tmp.Fields!dia23 + lnD23
                        mrs_Tmp.Fields!dia24 = mrs_Tmp.Fields!dia24 + lnD24
                        mrs_Tmp.Fields!dia25 = mrs_Tmp.Fields!dia25 + lnD25
                        mrs_Tmp.Fields!dia26 = mrs_Tmp.Fields!dia26 + lnD26
                        mrs_Tmp.Fields!dia27 = mrs_Tmp.Fields!dia27 + lnD27
                        mrs_Tmp.Fields!dia28 = mrs_Tmp.Fields!dia28 + lnD28
                        mrs_Tmp.Fields!dia29 = mrs_Tmp.Fields!dia29 + lnD29
                        mrs_Tmp.Fields!dia30 = mrs_Tmp.Fields!dia30 + lnD30
                        mrs_Tmp.Fields!dia31 = mrs_Tmp.Fields!dia31 + lnD31
                        mrs_Tmp.Fields!Total = mrs_Tmp.Fields!Total + 1
                        mrs_Tmp.Update
                     End If
                End If
                '
                mrs_Tmp1.MoveNext
                If mrs_Tmp1.EOF Then
                   Exit Do
                End If
            Loop
            
        Loop
        
        lcPie = ""
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
        Else
            mrs_Cab.Close
            mrs_Det.Close
            lcTexto1 = " SHAPE {select NroCuenta,idServicio,ServicioHosp from reporte_cabecera} AS padre" & _
                       " APPEND ({select NroCuenta,DevolucionHC," & _
                       "                 dia1,dia2,dia3,dia4,dia5," & _
                       "                 dia6,dia7,dia8,dia9,dia10," & _
                       "                 dia11,dia12,dia13,dia14,dia15," & _
                       "                 dia16,dia17,dia18,dia19,dia20," & _
                       "                 dia21,dia22,dia23,dia24,dia25," & _
                       "                 dia26,dia27,dia28,dia29,dia30," & _
                       "                 dia31,total" & _
                       "   FROM reporte_detalle} AS hijo" & _
                       " RELATE NroCuenta TO NroCuenta) AS hijo"
            mrs_Shape.Open lcTexto1, sighentidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
            With RpHFormatMes
                 .Orientation = rptOrientLandscape
                 .Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                 .Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                 Set .Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                 .Sections("pie").Controls("lblPie").Caption = lcPie
                 Set .DataSource = mrs_Shape
                 .DataMember = ""
                 With .Sections("CabGrupo")
                      .Controls("txtServicio").DataMember = ""
                      .Controls("txtServicio").DataField = "ServicioHosp"
                 End With
                 With .Sections("DetGrupo")
                      .Controls("txtHoras").DataMember = "Hijo"
                      .Controls("txtHoras").DataField = "DevolucionHC"
                      .Controls("Texto1").DataMember = "Hijo"
                      .Controls("Texto1").DataField = "dia1"
                      .Controls("Texto2").DataMember = "Hijo"
                      .Controls("Texto2").DataField = "dia2"
                      .Controls("Texto3").DataMember = "Hijo"
                      .Controls("Texto3").DataField = "dia3"
                      .Controls("Texto4").DataMember = "Hijo"
                      .Controls("Texto4").DataField = "dia4"
                      .Controls("Texto5").DataMember = "Hijo"
                      .Controls("Texto5").DataField = "dia5"
                      .Controls("Texto6").DataMember = "Hijo"
                      .Controls("Texto6").DataField = "dia6"
                      .Controls("Texto7").DataMember = "Hijo"
                      .Controls("Texto7").DataField = "dia7"
                      .Controls("Texto8").DataMember = "Hijo"
                      .Controls("Texto8").DataField = "dia8"
                      .Controls("Texto9").DataMember = "Hijo"
                      .Controls("Texto9").DataField = "dia9"
                      .Controls("Texto10").DataMember = "Hijo"
                      .Controls("Texto10").DataField = "dia10"
                      .Controls("Texto11").DataMember = "Hijo"
                      .Controls("Texto11").DataField = "dia11"
                      .Controls("Texto12").DataMember = "Hijo"
                      .Controls("Texto12").DataField = "dia12"
                      .Controls("Texto13").DataMember = "Hijo"
                      .Controls("Texto13").DataField = "dia13"
                      .Controls("Texto14").DataMember = "Hijo"
                      .Controls("Texto14").DataField = "dia14"
                      .Controls("Texto15").DataMember = "Hijo"
                      .Controls("Texto15").DataField = "dia15"
                      .Controls("Texto16").DataMember = "Hijo"
                      .Controls("Texto16").DataField = "dia16"
                      .Controls("Texto17").DataMember = "Hijo"
                      .Controls("Texto17").DataField = "dia17"
                      .Controls("Texto18").DataMember = "Hijo"
                      .Controls("Texto18").DataField = "dia18"
                      .Controls("Texto19").DataMember = "Hijo"
                      .Controls("Texto19").DataField = "dia19"
                      .Controls("Texto20").DataMember = "Hijo"
                      .Controls("Texto20").DataField = "dia20"
                      .Controls("Texto21").DataMember = "Hijo"
                      .Controls("Texto21").DataField = "dia21"
                      .Controls("Texto22").DataMember = "Hijo"
                      .Controls("Texto22").DataField = "dia22"
                      .Controls("Texto23").DataMember = "Hijo"
                      .Controls("Texto23").DataField = "dia23"
                      .Controls("Texto24").DataMember = "Hijo"
                      .Controls("Texto24").DataField = "dia24"
                      .Controls("Texto25").DataMember = "Hijo"
                      .Controls("Texto25").DataField = "dia25"
                      .Controls("Texto26").DataMember = "Hijo"
                      .Controls("Texto26").DataField = "dia26"
                      .Controls("Texto27").DataMember = "Hijo"
                      .Controls("Texto27").DataField = "dia27"
                      .Controls("Texto28").DataMember = "Hijo"
                      .Controls("Texto28").DataField = "dia28"
                      .Controls("Texto29").DataMember = "Hijo"
                      .Controls("Texto29").DataField = "dia29"
                      .Controls("Texto30").DataMember = "Hijo"
                      .Controls("Texto30").DataField = "dia30"
                      .Controls("Texto31").DataMember = "Hijo"
                      .Controls("Texto31").DataField = "dia31"
                      .Controls("Texto32").DataMember = "Hijo"
                      .Controls("Texto32").DataField = "total"
                 End With
                 .Show 1
            End With
            'debb-27/05/2015
            Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
            mo_reglasComunes.grabaTablaAuditoria ("RpHFormatMes: " & Mid(lcSubTitulo, IIf(InStr(lcSubTitulo, "FILTROS: ") > 0, 10, 1)))
            Set mo_reglasComunes = Nothing
            '
        End If
    End If
End Sub

