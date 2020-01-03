VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form MonitoresTV 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   4380
   ClientWidth     =   13560
   Icon            =   "MonitoresTV.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   13560
   Begin VB.Frame FraCuposLibres 
      Height          =   7785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13560
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   4905
         Top             =   90
      End
      Begin UltraGrid.SSUltraGrid grdCupos2 
         Height          =   6855
         Left            =   6960
         TabIndex        =   2
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   12091
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdCupos2"
      End
      Begin UltraGrid.SSUltraGrid grdCupos1 
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   12938
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdCupos1"
      End
      Begin VB.Image pi_ImagSeleccionada 
         BorderStyle     =   1  'Fixed Single
         Height          =   1170
         Left            =   0
         MouseIcon       =   "MonitoresTV.frx":000C
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label lblTextoCabecera 
         Alignment       =   2  'Center
         Caption         =   "Cupos Hasta : XX/XX/XXXXX xx:xx:xx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   6960
         TabIndex        =   3
         Top             =   120
         Width           =   6375
      End
   End
End
Attribute VB_Name = "MonitoresTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnAnchoPantalla As Long: Dim lnLargoPantalla As Long
Dim oRsServiciosCuposLibre1 As New Recordset
Dim oRsServiciosCuposLibre2 As New Recordset
Dim lbTodaviaProcesando As Boolean
Dim LnTotalRegistrosGrilla1 As Integer
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasComunes  As New SIGHNegocios.ReglasComunes
Const LnWidthFrame = 13575
Const LnTopFrame = 360
Const LnLeftFrame = 120
Const LnHeightFrame = 7455

Const LxCeroPaciente As String = "TOCA ATENCION"
Const LxPasoHoraAtencion As String = "PASO SU HORA DE ATENCION"
Dim lbMuestraImagen As Boolean, lcRutaImg As String

Private Sub Form_Activate()
    Me.Top = 0
    Me.Left = 0
    Me.Width = lnAnchoPantalla
    Me.Height = lnLargoPantalla
End Sub

Private Sub Form_Load()
    On Error GoTo ErrLoad
    
    lcRutaImg = App.Path & "\imagen.jpg"
    If SIGHEntidades.ArchivoExiste(lcRutaImg) Then
       pi_ImagSeleccionada.Picture = LoadPicture(lcRutaImg)
       lbMuestraImagen = True
    Else
       lbMuestraImagen = False
    End If
    
    LnTotalRegistrosGrilla1 = Val(lcBuscaParametro.SeleccionaFilaParametro(311))
    
    grdCupos1.Caption = ""
    grdCupos2.Caption = ""
    
    lnAnchoPantalla = Screen.Width
    lnLargoPantalla = Screen.Height

    
    If wxMuestraGrid = "CuposLibres" Then
        Me.Caption = "SisGalenPlus - CUPOS LIBRES en Admisión de Citas"
        FraCuposLibres.Width = LnWidthFrame
        FraCuposLibres.Top = LnTopFrame
        FraCuposLibres.Left = LnLeftFrame
        FraCuposLibres.Height = LnHeightFrame
        
        CreaTemporalCuposLibres1
        CreaTemporalCuposLibres2
    ElseIf wxMuestraGrid = "ATENCIONCE" Then
        Me.Caption = "SisGalenPlus - Pacientes atendidos o por atenderse en Consultorios Externos"
        FraCuposLibres.Top = 0
        FraCuposLibres.Left = 0
        FraCuposLibres.Width = Screen.Width - 300
        FraCuposLibres.Height = Screen.Height - 500
        lblTextoCabecera.Visible = False
        grdCupos2.Visible = False
        grdCupos1.Top = FraCuposLibres.Top
        grdCupos1.Width = FraCuposLibres.Width
        grdCupos1.Left = FraCuposLibres.Left
        grdCupos1.Height = FraCuposLibres.Height
        lbTodaviaProcesando = True
    End If
    Exit Sub
ErrLoad:
    MsgBox Err.Description
End Sub
Sub CreaTemporalCuposLibres1()
    If oRsServiciosCuposLibre1.State = 1 Then
       Set oRsServiciosCuposLibre1 = Nothing
    End If
    With oRsServiciosCuposLibre1
          .Fields.Append "IdServicio", adInteger
          .Fields.Append "Servicio", adVarChar, 255, adFldIsNullable
          .Fields.Append "Turno", adVarChar, 255, adFldIsNullable
          .Fields.Append "CuposLibres", adVarChar, 255, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdCupos1.DataSource = oRsServiciosCuposLibre1
    mo_Apariencia.ConfigurarFilasBiColores grdCupos1, SIGHEntidades.GrillaConFilasBicolor
End Sub
Sub CreaTemporalCuposLibres2()
    If oRsServiciosCuposLibre2.State = 1 Then
       Set oRsServiciosCuposLibre2 = Nothing
    End If
    With oRsServiciosCuposLibre2
          .Fields.Append "IdServicio", adInteger
          .Fields.Append "Servicio", adVarChar, 255, adFldIsNullable
          .Fields.Append "Turno", adVarChar, 255, adFldIsNullable
          .Fields.Append "CuposLibres", adVarChar, 255, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdCupos2.DataSource = oRsServiciosCuposLibre2
    mo_Apariencia.ConfigurarFilasBiColores grdCupos2, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub Timer1_Timer()
   Select Case wxMuestraGrid
   Case "ATENCIONCE"
        MuestraPacientesCitadosEnConsultorios
   Case "CuposLibres"
        MuestraCuposLibres
   End Select
End Sub


Sub MuestraCuposLibres()
    On Error GoTo ErrCerrar
    Dim ldFechaActual As Date
    Dim ldHoraActual As String
    ldFechaActual = Date
    ldHoraActual = Format$(Now, "h:mm")
    If wxMuestraGrid = "CuposLibres" Then
        'Configura ventana al tamaño maximo
        lblTextoCabecera.Width = Screen.Width - 300
        FraCuposLibres.Top = 0
        FraCuposLibres.Width = Screen.Width - 300
        FraCuposLibres.Height = Screen.Height - 500   'Screen.Height - 1900
        grdCupos1.Left = 100
        grdCupos1.Width = (Screen.Width - 300) / 2 - 200
        grdCupos1.Height = Screen.Height - 600
        
        grdCupos1.Bands(0).Columns("Servicio").Width = ((Screen.Width - 300) / 2 - 200) / 2 - 150
        grdCupos1.Bands(0).Columns("Turno").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        grdCupos1.Bands(0).Columns("CuposLibres").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        
        grdCupos2.Left = 200 + (Screen.Width - 300) / 2 - 200
        grdCupos2.Width = (Screen.Width - 300) / 2 - 200
        grdCupos2.Height = Screen.Height - 600
        
        grdCupos2.Bands(0).Columns("Servicio").Width = ((Screen.Width - 300) / 2 - 200) / 2 - 150
        grdCupos2.Bands(0).Columns("Turno").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        grdCupos2.Bands(0).Columns("CuposLibres").Width = ((Screen.Width - 300) / 2 - 200) / 4 - 150
        
   
    
        Dim oRsTmpProgMedServicios As New Recordset
        Dim oRsTmpCitas As New Recordset
        Dim lcFecha As String
        
        Dim lcHoraLimite As String
        
        Dim lHoraInicio As Long
        Dim lHoraFin  As Long
        Dim lHoraActual As Long
        Dim lTiempoPromedio As Long
        Dim lHoraSiguiente As Long
        Dim lHoraLimite As Long
        Dim lnTotalCupos As Integer, lnIdTurno As Long
        Dim lnTotalCuposBloqueados As Integer
        Dim lcHoraInicio As String, lcHoraFinal As String
        Dim lnTotalCitas As Integer
        Dim lcTextoTotalCupos As String
        Dim lnRegistroGrdCupos As Integer
        Dim lbEsHospitalTarapoto As Boolean
        lbEsHospitalTarapoto = False

        
        
        LimpiarTemporalesCuposLibres
        lnRegistroGrdCupos = 0
        


'    LnTotalRegistrosGrilla1
        lblTextoCabecera.Caption = "Cupos desde: " & ldFechaActual & " " & ldHoraActual
        Set oRsTmpProgMedServicios = mo_ReglasDeProgMedica.ProgramacionMedicaServiciosSeleccionarPorFechas(ldFechaActual, ldFechaActual)
        If oRsTmpProgMedServicios.RecordCount <= 0 Then
            oRsTmpProgMedServicios.Close
            Set oRsTmpProgMedServicios = Nothing
        End If
        oRsTmpProgMedServicios.MoveFirst
        Do While Not oRsTmpProgMedServicios.EOF
            'Calcula Total Cupos
            lHoraInicio = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmpProgMedServicios.Fields!HoraInicio)
            lHoraFin = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmpProgMedServicios.Fields!HoraFin)
            lHoraActual = mo_ReglasDeProgMedica.ConvertirAMinutos(ldHoraActual)
            
            If lHoraActual <= lHoraFin Then
                lTiempoPromedio = oRsTmpProgMedServicios.Fields!TiempoPromedioAtencion
                lHoraSiguiente = lHoraInicio
                lnTotalCupos = 0
                lnTotalCuposBloqueados = 0
                lHoraLimite = 0
                Do While lHoraSiguiente < lHoraFin
                    lnTotalCupos = lnTotalCupos + 1
                    If lHoraLimite = 0 Then
                        If lHoraSiguiente <= lHoraActual And lHoraActual <= lHoraSiguiente + lTiempoPromedio Then
                            lHoraLimite = lHoraSiguiente
                            If lHoraActual = lHoraSiguiente + lTiempoPromedio Then
                                lHoraLimite = lHoraSiguiente + lTiempoPromedio
                            End If
                        End If
                    End If
                    If lHoraSiguiente + lTiempoPromedio <= lHoraActual Then
                        lnTotalCuposBloqueados = lnTotalCuposBloqueados + 1
                    End If
                    lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
'                    lcHoraInicio = mo_ReglasDeProgMedica.ConvertirAHora(lHoraInicio)
'                    lcHoraFinal = mo_ReglasDeProgMedica.ConvertirAHora(lHoraSiguiente)
'                    lHoraInicio = lHoraSiguiente
                Loop
                lcHoraLimite = mo_ReglasDeProgMedica.ConvertirAHora(lHoraLimite)
                If lHoraLimite = 0 Then lcHoraLimite = mo_ReglasDeProgMedica.ConvertirAHora(lHoraInicio)
                
                
                Set oRsTmpCitas = mo_ReglasDeProgMedica.CitasSeleccionarPorServicioTurnoFecha(ldFechaActual, oRsTmpProgMedServicios.Fields!IdServicio, oRsTmpProgMedServicios.Fields!IdTurno, lcHoraLimite)
                lnTotalCitas = oRsTmpCitas.RecordCount
                
                If lnRegistroGrdCupos < LnTotalRegistrosGrilla1 Then
                    'Agregar Informacion Cupos
                    oRsServiciosCuposLibre1.AddNew
                    oRsServiciosCuposLibre1.Fields!IdServicio = oRsTmpProgMedServicios.Fields!IdServicio
                    oRsServiciosCuposLibre1.Fields!servicio = oRsTmpProgMedServicios.Fields!servicio
                    oRsServiciosCuposLibre1.Fields!Turno = oRsTmpProgMedServicios.Fields!HoraInicio & " - " & oRsTmpProgMedServicios.Fields!HoraFin
                    
                    lcTextoTotalCupos = "No hay"
                    If lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas > 0 Then
                        lcTextoTotalCupos = lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas
                    End If
                    oRsServiciosCuposLibre1.Fields!CuposLibres = lcTextoTotalCupos
                    oRsServiciosCuposLibre1.Update
                    lnRegistroGrdCupos = lnRegistroGrdCupos + 1 'Cuenta Registro
                Else
                    'Agregar Informacion Cupos
                    oRsServiciosCuposLibre2.AddNew
                    oRsServiciosCuposLibre2.Fields!IdServicio = oRsTmpProgMedServicios.Fields!IdServicio
                    oRsServiciosCuposLibre2.Fields!servicio = oRsTmpProgMedServicios.Fields!servicio
                    oRsServiciosCuposLibre2.Fields!Turno = oRsTmpProgMedServicios.Fields!HoraInicio & " - " & oRsTmpProgMedServicios.Fields!HoraFin
                    
                    lcTextoTotalCupos = "No hay"
                    If lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas > 0 Then
                        lcTextoTotalCupos = lnTotalCupos - lnTotalCuposBloqueados - lnTotalCitas
                    End If
                    oRsServiciosCuposLibre2.Fields!CuposLibres = lcTextoTotalCupos
                    oRsServiciosCuposLibre2.Update
                    lnRegistroGrdCupos = lnRegistroGrdCupos + 1 'Cuenta Registro
                End If
            End If
            oRsTmpProgMedServicios.MoveNext
        Loop
       
       Dim Row As SSRow
       
       With oRsServiciosCuposLibre1
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                   Set Row = Me.grdCupos1.ActiveRow
'                   row.Cells(3).Appearance.Font.Bold = True
                   If Row.Cells(3) = "No hay" Then
                        Row.Cells(3).Appearance.ForeColor = &HFF&
                   End If
                  .MoveNext
               Loop
            End If
        End With
        
       With oRsServiciosCuposLibre2
            If .RecordCount > 0 Then
               .MoveFirst
               Do While Not .EOF
                   Set Row = Me.grdCupos2.ActiveRow
'                   row.Cells(3).Appearance.Font.Bold = True
                   If Row.Cells(3) = "No hay" Then
                        Row.Cells(3).Appearance.ForeColor = &HFF&
                   End If
                  .MoveNext
               Loop
            End If
        End With
    End If
    Exit Sub
ErrCerrar:
    If Err.Number = 3705 Then
'      Select Case LnGrid
'      Case 1
'          oRsHospitalizados.Close
'          Resume
'      End Select
    Else
       MsgBox Err.Description
    End If
    Me.MousePointer = 1
    Exit Sub
    Resume

End Sub
Sub MuestraPacientesCitadosEnConsultorios()
    On Error GoTo ErrMesPac
    If lbTodaviaProcesando = True Then
        Dim ldFechaActual As Date, lnNumeroActual As Integer, lcQuedan As String, lnQuedan As Integer
        Dim lnIdAtencionUltimoAtendido As Long, lbYaAtendido As Boolean, lbDespuesDelUltimoAtendido As Boolean
        Dim lnHoraIngresoUltimoAtendido As String
        Dim lcPasoTriaje As String
        Dim oRsPacientesCitados As New Recordset
        Dim oRsPacientesSinAtender As New Recordset
        Dim oRsPacientesCitadosYAtendidos As New Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim lcConsultorios As String
        Dim oConexionExterna As New Connection
        oConexionExterna.CommandTimeout = 900
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        Set oRsTmp2 = mo_ReglasFacturacion.ServiciosSeleccionarPorFiltro("idTipoServicio=1", sghPorCodigo)
        If InStr(wxReniecHoraFin, "/") = 0 Then
           lcConsultorios = ""
        Else
           lcConsultorios = wxReniecHoraFin
        End If
        ldFechaActual = Date
        'ldFechaActual = CDate("15/01/2016")
        Do While True
            Set oRsPacientesCitados = mo_AdminAdmision.AtencionesCEseleccionarCITADOS(ldFechaActual)
            If oRsPacientesCitados.RecordCount > 0 Then
               CreaTmpPacientesCitados oRsPacientesSinAtender
               lnNumeroActual = 1
               oRsPacientesCitados.MoveFirst
               Do While Not oRsPacientesCitados.EOF
                  lbYaAtendido = False
                  If lcConsultorios <> "" Then
                     If InStr(lcConsultorios, "/" & Trim(Str(oRsPacientesCitados!IdServicioIngreso)) & "/") = 0 Then
                        lbYaAtendido = True
                     End If
                  End If
                  lcQuedan = "": lnIdAtencionUltimoAtendido = 0
                  Set oRsPacientesCitadosYAtendidos = mo_AdminAdmision.AtencionesCEseleccionarCitadosYAtendidos(ldFechaActual, _
                                                      oRsPacientesCitados!IdServicioIngreso)
                  If lbYaAtendido = False Then
                        oRsPacientesCitadosYAtendidos.Filter = "idAtencion=" & oRsPacientesCitados!idAtencion
                        If oRsPacientesCitadosYAtendidos.RecordCount > 0 Then
                           If Not IsNull(oRsPacientesCitadosYAtendidos!HoraEgreso) Then
                              If wxReniecHoraInicio = "*" Then
                                 lbYaAtendido = True
                              Else
                                 lcQuedan = "YA ATENDIDO"
                              End If
                           End If
                        End If
                  End If
                  If lbYaAtendido = False And lcQuedan = "" Then
                        lnQuedan = 0
                        oRsPacientesCitadosYAtendidos.Filter = ""
                        If oRsPacientesCitadosYAtendidos.RecordCount > 0 Then
                           oRsPacientesCitadosYAtendidos.MoveFirst
                           If IsNull(oRsPacientesCitadosYAtendidos!HoraEgreso) Then
                                lnIdAtencionUltimoAtendido = 0
                                lnHoraIngresoUltimoAtendido = ""
                                lbDespuesDelUltimoAtendido = True
                           Else
                                lnIdAtencionUltimoAtendido = oRsPacientesCitadosYAtendidos!idAtencion
                                lnHoraIngresoUltimoAtendido = oRsPacientesCitadosYAtendidos!HoraIngreso
                                lbDespuesDelUltimoAtendido = False
                           End If
                           oRsPacientesCitadosYAtendidos.Sort = "horaIngreso"
                           oRsPacientesCitadosYAtendidos.MoveFirst
                           If lnIdAtencionUltimoAtendido > 0 Then
                              oRsPacientesCitadosYAtendidos.Find "idAtencion=" & lnIdAtencionUltimoAtendido
                              'oRsPacientesCitadosYAtendidos.MoveNext
                           End If
                           If oRsPacientesCitadosYAtendidos.EOF Then
                                lcQuedan = LxPasoHoraAtencion
                           Else
                                
                                Do While Not oRsPacientesCitadosYAtendidos.EOF
                                   If lnIdAtencionUltimoAtendido = oRsPacientesCitadosYAtendidos!idAtencion Then
                                      lnQuedan = -1
                                      lbDespuesDelUltimoAtendido = True
                                   End If
                                   If oRsPacientesCitados!idAtencion = oRsPacientesCitadosYAtendidos!idAtencion Then
                                      If lbDespuesDelUltimoAtendido = True Then
                                            If lnQuedan = 0 Then
                                               lcQuedan = LxCeroPaciente
                                            Else
                                               lcQuedan = Trim(Str(lnQuedan)) & " PACIENTE" & IIf(lnQuedan = 1, "", "S")
                                            End If
                                      Else
                                            lcQuedan = LxPasoHoraAtencion
                                      End If
                                      Exit Do
                                   ElseIf Not IsNull(oRsPacientesCitadosYAtendidos!HoraEgreso) And lnIdAtencionUltimoAtendido <> oRsPacientesCitadosYAtendidos!idAtencion Then
                                      If oRsPacientesCitadosYAtendidos!HoraEgreso > lnHoraIngresoUltimoAtendido Then
                                         lnQuedan = lnQuedan - 1
                                      End If
                                   End If
                                   oRsPacientesCitadosYAtendidos.MoveNext
                                   lnQuedan = lnQuedan + 1
                                Loop
                                If lcQuedan = "" Then
                                   lcQuedan = LxPasoHoraAtencion
                                End If
                           End If
                        End If
                  End If
                  oRsPacientesCitadosYAtendidos.Close
                  If lbYaAtendido = False Then
                        'pasó por Triaje
                        lcPasoTriaje = "No necesario"
                        oRsTmp2.Filter = "triaje=1 and idServicio=" & oRsPacientesCitados!IdServicioIngreso
                        If oRsTmp2.RecordCount > 0 Then
                            Set oRsTmp1 = mo_AdminAdmision.atencionesCExServicio(oRsPacientesCitados!IdServicioIngreso, ldFechaActual, oConexionExterna)
                            oRsTmp1.Filter = "idAtencion=" & oRsPacientesCitados!idAtencion
                            lcPasoTriaje = "No"
                            If oRsTmp1.RecordCount > 0 Then
                               If (Not IsNull(oRsTmp1.Fields!TriajeFecha)) Then
                                   lcPasoTriaje = "Si"
                               End If
                            End If
                            oRsTmp1.Close
                        End If
                        '
                        oRsPacientesSinAtender.AddNew
                        oRsPacientesSinAtender!Paciente = oRsPacientesCitados!ApellidoPaterno & " " & oRsPacientesCitados!ApellidoMaterno & _
                                                          " " & oRsPacientesCitados!PrimerNombre & " " & IIf(IsNull(oRsPacientesCitados!SegundoNombre), "", oRsPacientesCitados!SegundoNombre)
                        oRsPacientesSinAtender!NroHistoria = SIGHEntidades.HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsPacientesCitados!NroHistoriaClinica)), False)
                        oRsPacientesSinAtender!Consultorio = oRsPacientesCitados!Consultorio
                        oRsPacientesSinAtender!Triaje = lcPasoTriaje
                        oRsPacientesSinAtender!quedan = lcQuedan
                        oRsPacientesSinAtender.Update
                        lnNumeroActual = lnNumeroActual + 1
                  End If
                  oRsPacientesCitados.MoveNext
                  If lnNumeroActual = wxNumMinutosGrid Then
                     lnNumeroActual = 1
                     grdCupos1.Caption = "CITADOS POR ATENDER: " & ldFechaActual
                     Set grdCupos1.DataSource = oRsPacientesSinAtender
                     mo_ReglasComunes.WaitSeconds 10
                     CreaTmpPacientesCitados oRsPacientesSinAtender
                  End If
               Loop
               Set grdCupos1.DataSource = oRsPacientesSinAtender
               mo_ReglasComunes.WaitSeconds 10
            End If
        
            'imagen por 30 segundos
            If lbMuestraImagen = True Then
                grdCupos1.Visible = False
                pi_ImagSeleccionada.Visible = True
                pi_ImagSeleccionada.Width = grdCupos1.Width
                pi_ImagSeleccionada.Height = grdCupos1.Height
                
                mo_ReglasComunes.WaitSeconds 30
                pi_ImagSeleccionada.Visible = False
                grdCupos1.Visible = True
            End If
            '
        Loop
        Set oRsPacientesCitados = Nothing
        Set oRsPacientesSinAtender = Nothing
        Set oRsPacientesCitadosYAtendidos = Nothing
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        lbTodaviaProcesando = False
    
    End If
    Exit Sub
ErrMesPac:
    grdCupos1.Visible = True
    pi_ImagSeleccionada.Visible = False
    Set oRsPacientesCitados = Nothing
    Set oRsPacientesSinAtender = Nothing
    Set oRsPacientesCitadosYAtendidos = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    lbTodaviaProcesando = False
End Sub

Sub LimpiarTemporalesCuposLibres()

    With oRsServiciosCuposLibre1
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
    
    With oRsServiciosCuposLibre2
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With

End Sub

Sub CreaTmpPacientesCitados(oRsPacientesSinAtender As Recordset)
    If oRsPacientesSinAtender.State = 1 Then oRsPacientesSinAtender.Close
    With oRsPacientesSinAtender
          .Fields.Append "Paciente", adVarChar, 100, adFldIsNullable
          .Fields.Append "NroHistoria", adVarChar, 20, adFldIsNullable
          .Fields.Append "Consultorio", adVarChar, 50, adFldIsNullable
          .Fields.Append "Triaje", adVarChar, 20, adFldIsNullable
          .Fields.Append "Quedan", adVarChar, 30, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    mo_Apariencia.ConfigurarFilasBiColores grdCupos1, SIGHEntidades.GrillaConFilasBicolor
End Sub
Private Sub grdCupos1_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    If wxMuestraGrid = "ATENCIONCE" Then
        If InStr(wxSisAcreditacioHoraInicio, ":") = 0 And Val(wxSisAcreditacioHoraInicio) > 0 Then
           grdCupos1.Bands(0).Columns("consultorio").Width = Val(wxSisAcreditacioHoraInicio)
        Else
           grdCupos1.Bands(0).Columns("consultorio").Width = 4000
        End If
        If InStr(wxSisAcreditacioHoraFinal, ":") = 0 And Val(wxSisAcreditacioHoraFinal) > 0 Then
           grdCupos1.Bands(0).Columns("Paciente").Width = Val(wxSisAcreditacioHoraFinal)
        Else
           grdCupos1.Bands(0).Columns("Paciente").Width = 7700
        End If
        
        grdCupos1.Bands(0).Columns("nroHistoria").Width = 2100
        grdCupos1.Bands(0).Columns("Quedan").Width = 3500
        grdCupos1.Bands(0).Columns("Triaje").Width = 2400
        'grdCupos1.Bands(0).Columns("idServicioIngreso").Hidden = True
        

        
    Else
        grdCupos1.Bands(0).Columns("IdServicio").Hidden = True
        grdCupos1.Bands(0).Columns("Servicio").Header.Caption = "Consultorios"
        grdCupos1.Bands(0).Columns("Turno").Header.Caption = "Turno"
        grdCupos1.Bands(0).Columns("CuposLibres").Header.Caption = "Cupos Libres"
        grdCupos1.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
        grdCupos1.Bands(0).Columns("Turno").Activation = ssActivationActivateNoEdit
        grdCupos1.Bands(0).Columns("CuposLibres").Activation = ssActivationActivateNoEdit
        grdCupos1.Bands(0).Columns("CuposLibres").CellAppearance.TextAlign = ssAlignCenter
    End If
End Sub



Private Sub grdCupos2_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCupos2.Bands(0).Columns("IdServicio").Hidden = True
    grdCupos2.Bands(0).Columns("Servicio").Header.Caption = "Consultorios"
    grdCupos2.Bands(0).Columns("Turno").Header.Caption = "Turno"
    grdCupos2.Bands(0).Columns("CuposLibres").Header.Caption = "Cupos Libres"
    grdCupos2.Bands(0).Columns("Servicio").Activation = ssActivationActivateNoEdit
    grdCupos2.Bands(0).Columns("Turno").Activation = ssActivationActivateNoEdit
    grdCupos2.Bands(0).Columns("CuposLibres").Activation = ssActivationActivateNoEdit
    grdCupos2.Bands(0).Columns("CuposLibres").CellAppearance.TextAlign = ssAlignCenter
End Sub


Private Sub grdCupos1_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If wxMuestraGrid = "ATENCIONCE" Then
        Select Case Row.Cells("Quedan").GetText()
        Case LxCeroPaciente
            Row.Appearance.ForeColor = vbRed
            
        End Select
    End If
End Sub

