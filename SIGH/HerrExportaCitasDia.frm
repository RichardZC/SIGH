VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form HerrExportaCitasDia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programación médica de un Servicio en un día"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14415
   ControlBox      =   0   'False
   Icon            =   "HerrExportaCitasDia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8745
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   14265
      Begin VB.Frame FraFiltros 
         Height          =   2145
         Left            =   10905
         TabIndex        =   16
         Top             =   3975
         Width           =   3270
         Begin VB.TextBox txtPorcI 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   495
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "0"
            Top             =   1260
            Width           =   405
         End
         Begin VB.TextBox txtPorcS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "0"
            Top             =   735
            Width           =   405
         End
         Begin VB.CheckBox chkTodos 
            Caption         =   "Todos/Ninguno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   2
            Top             =   300
            Width           =   2025
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "% de los últimos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   945
            TabIndex        =   20
            Top             =   1320
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   285
            TabIndex        =   19
            Top             =   1320
            Width           =   150
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "% de los primeros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   930
            TabIndex        =   18
            Top             =   795
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   270
            TabIndex        =   17
            Top             =   795
            Width           =   150
         End
      End
      Begin VB.CommandButton cmdActualizaTotalCupos 
         Caption         =   "Totaliza Cupos Web"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10860
         TabIndex        =   15
         Top             =   3270
         Width           =   3315
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Exportar Citas Web"
         DisabledPicture =   "HerrExportaCitasDia.frx":0CCA
         DownPicture     =   "HerrExportaCitasDia.frx":112A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11160
         Picture         =   "HerrExportaCitasDia.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7860
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrExportaCitasDia.frx":1A14
         DownPicture     =   "HerrExportaCitasDia.frx":1ED8
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12630
         Picture         =   "HerrExportaCitasDia.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7860
         Width           =   1335
      End
      Begin VB.TextBox txtCuposWeb 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10830
         TabIndex        =   12
         Top             =   2880
         Width           =   3345
      End
      Begin VB.TextBox txtTotalCupos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10830
         TabIndex        =   11
         Top             =   2120
         Width           =   3345
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10830
         TabIndex        =   10
         Top             =   1360
         Width           =   3345
      End
      Begin VB.TextBox txtServicio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10830
         TabIndex        =   9
         Top             =   600
         Width           =   3345
      End
      Begin UltraGrid.SSUltraGrid grdProgramacionDelDia 
         Height          =   8310
         Left            =   120
         TabIndex        =   4
         Top             =   330
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   14658
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         RowConnectorColor=   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdAnteriores"
      End
      Begin VB.Label Label4 
         Caption         =   "Total cupos separados para Citas Web"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         TabIndex        =   8
         Top             =   2640
         Width           =   3195
      End
      Begin VB.Label Label3 
         Caption         =   "Total cupos programados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         TabIndex        =   7
         Top             =   1875
         Width           =   2265
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha programada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         TabIndex        =   6
         Top             =   1095
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         TabIndex        =   5
         Top             =   330
         Width           =   765
      End
   End
End
Attribute VB_Name = "HerrExportaCitasDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exporta Citas WEB, también importa
'        Programado por: Barrantes D
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim oRsProgramacionDelDia As New Recordset
Dim oRsCuposDelDia As New Recordset
Dim oRsCitaWebCupos As New Recordset
Dim lcSql As String
Dim ml_IdServicio As Long
Dim ml_Fecha As Date, lnCuposDisponibles As Integer
Dim mi_BotonPresionado As sghBotonDetallePresionado
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Set CuposDelDia(oValue As Recordset)
    Set oRsCuposDelDia = oValue
End Property

Property Get CuposDelDia() As Recordset
    Set CuposDelDia = oRsCuposDelDia.Clone()
End Property

Property Let IdServicio(lIdValue As Long)
    ml_IdServicio = lIdValue
End Property

Property Let fecha(lIdValue As Date)
    ml_Fecha = lIdValue
End Property

Property Get NroCuposElegidos() As Long
    NroCuposElegidos = Val(Me.txtCuposWeb.Text)
End Property



Private Sub btnAceptar_Click()
        TotalizaCuposElegidos
        Dim lnTotalRegistros As Long, lbNuevo As Boolean
        lnTotalRegistros = oRsCuposDelDia.RecordCount
        If oRsProgramacionDelDia.RecordCount > 0 Then
           oRsProgramacionDelDia.MoveFirst
           Do While Not oRsProgramacionDelDia.EOF
              lbNuevo = True
              If lnTotalRegistros > 0 Then
                 oRsCuposDelDia.MoveFirst
                 oRsCuposDelDia.Find "HoraInicio='" & oRsProgramacionDelDia.Fields!HoraInicio & "'"
                 If Not oRsCuposDelDia.EOF Then
                    lbNuevo = False
                 End If
              End If
              If lbNuevo = True Then
                    oRsCuposDelDia.AddNew
                    oRsCuposDelDia.Fields!IdServicio = oRsProgramacionDelDia.Fields!IdServicio
                    oRsCuposDelDia.Fields!idMedico = oRsProgramacionDelDia.Fields!idMedico
                    oRsCuposDelDia.Fields!fecha = oRsProgramacionDelDia.Fields!fecha
                    oRsCuposDelDia.Fields!HoraInicio = oRsProgramacionDelDia.Fields!HoraInicio
                    oRsCuposDelDia.Fields!HoraFinal = oRsProgramacionDelDia.Fields!HoraFinal
              End If
              'MARIO
              If oRsProgramacionDelDia.Fields!Elegir = False Then
                 If IsNull(oRsProgramacionDelDia.Fields!idCitaBloqueada) Then
                    If oRsProgramacionDelDia!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
                      oRsCuposDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb
                    Else
                      oRsCuposDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos
                    End If
                 Else
                    oRsCuposDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos
                 End If
                 oRsCuposDelDia.Fields!ApellidoPaterno = oRsProgramacionDelDia.Fields!ApellidoPaterno
                 oRsCuposDelDia.Fields!ApellidoMaterno = oRsProgramacionDelDia.Fields!ApellidoMaterno
                 oRsCuposDelDia.Fields!PrimerNombre = oRsProgramacionDelDia.Fields!PrimerNombre
                 oRsCuposDelDia.Fields!SegundoNombre = oRsProgramacionDelDia.Fields!SegundoNombre
                 oRsCuposDelDia.Fields!FechaNacimiento = oRsProgramacionDelDia.Fields!FechaNacimiento
                 oRsCuposDelDia.Fields!DNI = oRsProgramacionDelDia.Fields!DNI
                 oRsCuposDelDia.Fields!idTipoSexo = oRsProgramacionDelDia.Fields!idTipoSexo
                 oRsCuposDelDia.Fields!Ubigeo = oRsProgramacionDelDia.Fields!Ubigeo
                 oRsCuposDelDia.Fields!idPaciente = oRsProgramacionDelDia.Fields!idPaciente
              Else
                 oRsCuposDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoDisponibleEnCitaWeb
              End If
              oRsCuposDelDia.Fields!IdTurno = oRsProgramacionDelDia.Fields!IdTurno
              oRsCuposDelDia.Update
              oRsProgramacionDelDia.MoveNext
           Loop
           mi_BotonPresionado = sghAceptar
        End If
        Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub



Private Sub chkTodos_Click()
    txtPorcI.Text = "0"
    txtPorcS.Text = "0"
    oRsProgramacionDelDia.MoveFirst
    Do While Not oRsProgramacionDelDia.EOF
        If chkTodos.Value = 1 Then
            If oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Or _
                    oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Or _
                    oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
               oRsProgramacionDelDia.Fields!Elegir = False
            Else
               oRsProgramacionDelDia.Fields!Elegir = True
            End If
        Else
            oRsProgramacionDelDia.Fields!Elegir = False
        End If
        oRsProgramacionDelDia.Update
        oRsProgramacionDelDia.MoveNext
    Loop
    TotalizaCuposElegidos
End Sub

Private Sub cmdActualizaTotalCupos_Click()
     TotalizaCuposElegidos
End Sub

Private Sub Form_Load()
       mi_BotonPresionado = sghCancelar
       CargaDatosDeCuposPorServicioYfecha
       If CDate(lcBuscaParametro.RetornaFechaServidorSQL) = ml_Fecha Then
          btnAceptar.Visible = False
          Me.FraFiltros.Visible = False
          Me.grdProgramacionDelDia.Enabled = False
       End If
End Sub

'debb-18/06/2019
Sub CargaDatosDeCuposPorServicioYfecha()
    Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
    Dim oRsTurnos As New Recordset
    Dim lHoraInicio As Long
    Dim lHoraFin  As Long
    Dim lTiempoPromedio As Long
    Dim lHoraSiguiente As Long, lcDNIcitaWeb As String
    Dim lnTotalCupos As Integer, lnIdTurno As Long
    Dim lcHoraInicio As String, lcHoraFinal As String, lbElegir As Boolean
    If oRsProgramacionDelDia.State = 1 Then Set oRsProgramacionDelDia = Nothing
    With oRsProgramacionDelDia
          .Fields.Append "Elegir", adBoolean
          .Fields.Append "idServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "Fecha", adDate
          .Fields.Append "HoraInicio", adVarChar, 5, adFldIsNullable
          .Fields.Append "HoraFinal", adVarChar, 5, adFldIsNullable
          .Fields.Append "Medico", adVarChar, 100, adFldIsNullable
          .Fields.Append "idMedico", adInteger, 4, adFldIsNullable
          .Fields.Append "ApellidoPaterno", adVarChar, 40, adFldIsNullable
          .Fields.Append "ApellidoMaterno", adVarChar, 40, adFldIsNullable
          .Fields.Append "PrimerNombre", adVarChar, 40, adFldIsNullable
          .Fields.Append "SegundoNombre", adVarChar, 20, adFldIsNullable
          .Fields.Append "FechaNacimiento", adDate, 8, adFldIsNullable
          .Fields.Append "DNI", adVarChar, 8, adFldIsNullable
          .Fields.Append "idTipoSexo", adInteger, 4, adFldIsNullable
          .Fields.Append "idEstadoCitaWeb", adInteger, 4, adFldIsNullable    '1->llenado en Cita, 2-> Disponible para Web, 3->Confirmada en Web
          .Fields.Append "idCitaBloqueada", adInteger, 4, adFldIsNullable
          .Fields.Append "Ubigeo", adInteger, 4
          .Fields.Append "idTurno", adInteger, 4
          .Fields.Append "idPaciente", adInteger, 4
          .Fields.Append "Numero", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdProgramacionDelDia.DataSource = oRsProgramacionDelDia
    mo_Apariencia.ConfigurarFilasBiColores Me.grdProgramacionDelDia, sighentidades.GrillaConFilasBicolor
    grdProgramacionDelDia.Caption = ""
    
    Set oRsTmp1 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorFechas(ml_Fecha, ml_Fecha)
    oRsTmp1.Filter = "idServicio=" & ml_IdServicio
    If oRsTmp1.RecordCount = 0 Then
       MsgBox "No existe información", vbInformation, Me.Caption
    Else
       Set oRsTurnos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
       Set oRsTmp2 = mo_ReglasDeProgMedica.CitasSeleccionarPorServicioYfecha(ml_IdServicio, ml_Fecha)
       oRsTmp1.MoveFirst
       Me.txtServicio.Text = oRsTmp1.Fields!nombre
       Me.txtFecha.Text = ml_Fecha
       lnTotalCupos = 0: lnCuposDisponibles = 0
       Do While Not oRsTmp1.EOF
            lHoraInicio = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmp1.Fields!HoraInicio)
            lHoraFin = mo_ReglasDeProgMedica.ConvertirAMinutos(oRsTmp1.Fields!HoraFin)
            lTiempoPromedio = oRsTmp1.Fields!TiempoPromedioAtencion
            lHoraSiguiente = lHoraInicio
            Do While lHoraSiguiente < lHoraFin
                lnTotalCupos = lnTotalCupos + 1
                lHoraSiguiente = lHoraSiguiente + lTiempoPromedio
                lcHoraInicio = mo_ReglasDeProgMedica.ConvertirAHora(lHoraInicio)
                lcHoraFinal = mo_ReglasDeProgMedica.ConvertirAHora(lHoraSiguiente)
                '
                lnIdTurno = 4
                If oRsTurnos.RecordCount > 0 Then
                   oRsTurnos.MoveFirst
                   Do While Not oRsTurnos.EOF
                      If lcHoraInicio >= oRsTurnos.Fields!HoraInicio And lcHoraInicio < oRsTurnos.Fields!HoraFin Then
                         lnIdTurno = oRsTurnos.Fields!IdTurno
                         Exit Do
                      End If
                      oRsTurnos.MoveNext
                   Loop
                End If
                '
                If oRsTmp2.RecordCount > 0 Then
                   oRsTmp2.MoveFirst
                   Do While Not oRsTmp2.EOF
                      If oRsTmp2.Fields!HoraInicio >= lcHoraInicio And oRsTmp2.Fields!HoraInicio < lcHoraFinal Then
                         Exit Do
                      End If
                      oRsTmp2.MoveNext
                   Loop
                End If
                '
                lcDNIcitaWeb = ""
                lbElegir = False
                If oRsTmp2.EOF Then
                    If oRsCuposDelDia.RecordCount > 0 Then
                        oRsCuposDelDia.MoveFirst
                        Do While Not oRsCuposDelDia.EOF
                            If oRsCuposDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoDisponibleEnCitaWeb And oRsCuposDelDia.Fields!HoraInicio = lcHoraInicio And oRsCuposDelDia.Fields!HoraFinal = lcHoraFinal Then
                               lbElegir = True
                            ElseIf oRsCuposDelDia!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb And oRsCuposDelDia.Fields!HoraInicio = lcHoraInicio And oRsCuposDelDia.Fields!HoraFinal = lcHoraFinal Then
                                lcDNIcitaWeb = oRsCuposDelDia!DNI
                            End If
                            oRsCuposDelDia.MoveNext
                        Loop
                    End If
                End If
                oRsProgramacionDelDia.AddNew
                oRsProgramacionDelDia.Fields!HoraInicio = lcHoraInicio
                oRsProgramacionDelDia.Fields!HoraFinal = lcHoraFinal
                oRsProgramacionDelDia.Fields!IdServicio = ml_IdServicio
                oRsProgramacionDelDia.Fields!fecha = ml_Fecha
                oRsProgramacionDelDia.Fields!Medico = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & oRsTmp1.Fields!Nombres
                oRsProgramacionDelDia.Fields!idMedico = oRsTmp1.Fields!idMedico
                oRsProgramacionDelDia.Fields!Elegir = lbElegir
                oRsProgramacionDelDia.Fields!IdTurno = lnIdTurno
                
                If Not oRsTmp2.EOF Then
'                    oRsProgramacionDelDia.Fields!IdEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos
                    Set oRsCitaWebCupos = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ml_Fecha, ml_Fecha, oRsTmp1.Fields!idMedico, ml_IdServicio)
                    If oRsCitaWebCupos.RecordCount > 0 Then
                        Do While Not oRsCitaWebCupos.EOF
                        If oRsCitaWebCupos.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos And oRsCitaWebCupos.Fields!HoraInicio = lcHoraInicio And oRsCitaWebCupos.Fields!HoraFinal = lcHoraFinal Then
                            oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos
                            oRsProgramacionDelDia.Fields!idCitaBloqueada = oRsCitaWebCupos.Fields!idCitaBloqueada
                            Exit Do
                        Else
                            oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos
                        End If
                        oRsCitaWebCupos.MoveNext
                        Loop
                        Set oRsCitaWebCupos = Nothing
                    Else
                        oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos
                    End If
                    
                    If oRsTmp2.Fields!IdDocIdentidad = 1 Then
                        oRsProgramacionDelDia.Fields!DNI = oRsTmp2.Fields!nrodocumento
                    End If
                    oRsProgramacionDelDia.Fields!ApellidoPaterno = oRsTmp2.Fields!ApellidoPaterno
                    oRsProgramacionDelDia.Fields!ApellidoMaterno = oRsTmp2.Fields!ApellidoMaterno
                    oRsProgramacionDelDia.Fields!PrimerNombre = oRsTmp2.Fields!PrimerNombre
                    oRsProgramacionDelDia.Fields!SegundoNombre = IIf(IsNull(oRsTmp2.Fields!SegundoNombre), "", oRsTmp2.Fields!SegundoNombre)
                    oRsProgramacionDelDia.Fields!idTipoSexo = oRsTmp2.Fields!idTipoSexo
                    oRsProgramacionDelDia.Fields!FechaNacimiento = oRsTmp2.Fields!FechaNacimiento
                    oRsProgramacionDelDia.Fields!Ubigeo = IIf(IsNull(oRsTmp2.Fields!IdDistritoDomicilio), 0, oRsTmp2.Fields!IdDistritoDomicilio)
                    oRsProgramacionDelDia.Fields!idPaciente = oRsTmp2.Fields!idPaciente
                ElseIf lcDNIcitaWeb <> "" Then
                   oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb
                   oRsProgramacionDelDia.Fields!DNI = lcDNIcitaWeb
                Else
                   lnCuposDisponibles = lnCuposDisponibles + 1
                   oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoDisponibleEnCitaWeb
                   oRsProgramacionDelDia.Fields!Ubigeo = 0
                   oRsProgramacionDelDia!numero = lnCuposDisponibles
                   
                End If
                oRsProgramacionDelDia.Update
                lHoraInicio = lHoraSiguiente
            Loop
            oRsTmp1.MoveNext
       Loop
       Me.txtTotalCupos.Text = lnTotalCupos
    End If
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsTurnos = Nothing
    mo_Formulario.HabilitarDeshabilitar txtServicio, False
    mo_Formulario.HabilitarDeshabilitar txtFecha, False
    mo_Formulario.HabilitarDeshabilitar txtTotalCupos, False
    mo_Formulario.HabilitarDeshabilitar txtCuposWeb, False
    cmdActualizaTotalCupos_Click
    If oRsProgramacionDelDia.RecordCount > 0 Then oRsProgramacionDelDia.MoveFirst
End Sub

Sub TotalizaCuposElegidos()
    On Error GoTo ErrPrg
    Dim lnTotalCuposWeb As Integer, oRsProductos As New Recordset
    Set oRsProductos = oRsProgramacionDelDia.Clone
    oRsProductos.MoveFirst
    lnTotalCuposWeb = 0
    Do While Not oRsProductos.EOF
       If oRsProductos.Fields!Elegir = True Then
          lnTotalCuposWeb = lnTotalCuposWeb + 1
       End If
       oRsProductos.MoveNext
    Loop
    txtCuposWeb.Text = lnTotalCuposWeb
ErrPrg:

End Sub
'debb-18/05/2019
Private Sub grdProgramacionDelDia_BeforeCellActivate(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Dim Row As SSRow
    On Error Resume Next
    Set Row = grdProgramacionDelDia.ActiveCell.Row

    oRsProgramacionDelDia.MoveFirst
    Do While Not oRsProgramacionDelDia.EOF
        If oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Or _
                    oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Or _
                    oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
           oRsProgramacionDelDia.Fields!Elegir = False
           Row.Cells("elegir").Activation = ssActivationDisabled
        End If
        oRsProgramacionDelDia.Update
        oRsProgramacionDelDia.MoveNext
    Loop
    
    Set Row = Nothing
End Sub
'debb-18/05/2019
Private Sub grdProgramacionDelDia_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
On Error Resume Next
    Dim oRow As SSRow
    Set oRow = grdProgramacionDelDia.ActiveCell.Row
    Select Case grdProgramacionDelDia.ActiveCell.Column.Key
    Case "Elegir"
         If oRow.Cells("idEstadoCitaWeb").Value = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Or _
            oRow.Cells("idEstadoCitaWeb").Value = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Or _
            oRow.Cells("idEstadoCitaWeb").Value = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
            
            oRow.Cells("elegir").Value = False
            oRow.Cells("elegir").Activation = ssActivationDisabled
         End If
    End Select
    Set oRow = Nothing
    
End Sub

Private Sub grdProgramacionDelDia_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdProgramacionDelDia_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdProgramacionDelDia.Bands(0).Columns("idServicio").Hidden = True
    grdProgramacionDelDia.Bands(0).Columns("Fecha").Hidden = True
    grdProgramacionDelDia.Bands(0).Columns("idMedico").Hidden = True
    grdProgramacionDelDia.Bands(0).Columns("Numero").Hidden = True
    '
    grdProgramacionDelDia.Bands(0).Columns("Elegir").Width = 500
    grdProgramacionDelDia.Bands(0).Columns("HoraInicio").Width = 1000
    grdProgramacionDelDia.Bands(0).Columns("HoraInicio").Activation = ssActivationActivateNoEdit
    grdProgramacionDelDia.Bands(0).Columns("HoraFinal").Width = 1000
    grdProgramacionDelDia.Bands(0).Columns("HoraFinal").Activation = ssActivationActivateNoEdit
    grdProgramacionDelDia.Bands(0).Columns("Medico").Width = 2500
    grdProgramacionDelDia.Bands(0).Columns("Medico").Activation = ssActivationActivateNoEdit
    grdProgramacionDelDia.Bands(0).Columns("idPaciente").Width = 1000
    grdProgramacionDelDia.Bands(0).Columns("idPaciente").Activation = ssActivationActivateNoEdit
    
End Sub
'MARIO
Private Sub grdProgramacionDelDia_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
    If Row.Cells("IdEstadoCitaWeb").Value = 1 Or Row.Cells("IdEstadoCitaWeb").Value = 4 Then
        Row.Cells("elegir").Activation = ssActivationDisabled
    End If
End Sub


Sub PorPorcentaje(lnPorcentaje)
    Dim lnItems As Integer
    lnItems = Round(lnCuposDisponibles * lnPorcentaje / 100, 0)
    If Val(txtPorcS.Text) > 0 Then
        oRsProgramacionDelDia.MoveFirst
        Do While Not oRsProgramacionDelDia.EOF
            If oRsProgramacionDelDia!numero > 0 And oRsProgramacionDelDia!numero <= lnItems Then
                If oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Or _
                        oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Or _
                        oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
                   oRsProgramacionDelDia.Fields!Elegir = False
                Else
                   oRsProgramacionDelDia.Fields!Elegir = True
                End If
            Else
                oRsProgramacionDelDia.Fields!Elegir = False
            End If
            oRsProgramacionDelDia.Update
            oRsProgramacionDelDia.MoveNext
        Loop
    Else
        lnItems = lnCuposDisponibles - lnItems
        oRsProgramacionDelDia.MoveLast
        Do While Not oRsProgramacionDelDia.BOF
            If oRsProgramacionDelDia!numero > 0 And oRsProgramacionDelDia!numero > lnItems Then
                If oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Or _
                        oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Or _
                        oRsProgramacionDelDia.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
                   oRsProgramacionDelDia.Fields!Elegir = False
                Else
                   oRsProgramacionDelDia.Fields!Elegir = True
                End If
            Else
                oRsProgramacionDelDia.Fields!Elegir = False
            End If
            oRsProgramacionDelDia.Update
            oRsProgramacionDelDia.MovePrevious
        Loop
    End If
    TotalizaCuposElegidos
End Sub



Private Sub txtPorcI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPorcI
End Sub

Private Sub txtPorcI_LostFocus()
    If Val(txtPorcI.Text) > 0 And Val(txtPorcI.Text) <= 100 Then
       txtPorcS.Text = "0"
       PorPorcentaje Val(txtPorcI.Text)
    End If
End Sub


Private Sub txtPorcS_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtPorcS
End Sub

Private Sub txtPorcS_LostFocus()
    If Val(txtPorcS.Text) > 0 And Val(txtPorcS.Text) <= 100 Then
       txtPorcI.Text = "0"
       PorPorcentaje Val(txtPorcS.Text)
    End If

End Sub




