VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form AHhcNOusadas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historias Clínicas con Problemas"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "AHhcNOusadas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7590
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   9195
      Begin VB.CheckBox chkDelDNIdefault 
         Caption         =   "Eliminar el DNI usado como DEFAULT antes de procesar"
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
         Left            =   3180
         Picture         =   "AHhcNOusadas.frx":0CCA
         TabIndex        =   24
         Top             =   3390
         Width           =   4860
      End
      Begin VB.TextBox txtDNIusadoDefault 
         Height          =   315
         Left            =   8100
         TabIndex        =   23
         Top             =   3375
         Width           =   885
      End
      Begin VB.Frame Frame 
         Height          =   1635
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   465
         Width           =   8400
         Begin VB.TextBox txtHcMaxLinea 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2085
            TabIndex        =   18
            Text            =   "7"
            Top             =   525
            Width           =   585
         End
         Begin VB.TextBox txtNlineas 
            Height          =   285
            Left            =   2085
            TabIndex        =   17
            Text            =   "300"
            Top             =   915
            Width           =   585
         End
         Begin VB.CheckBox chkOrden 
            Alignment       =   1  'Right Justify
            Caption         =   "El orden de HC vacias es de MENOR a MAYOR ?"
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
            Left            =   135
            Picture         =   "AHhcNOusadas.frx":0FDC
            TabIndex        =   16
            Top             =   1245
            Value           =   1  'Checked
            Width           =   4215
         End
         Begin VB.Label lblIdTipoHistoria 
            AutoSize        =   -1  'True
            Caption         =   "Consideración: El reporte tomará  el último N° Historia Clínica generada en forma automática"
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
            Left            =   135
            TabIndex        =   21
            Top             =   165
            Width           =   7515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "N° HC máximas x linea"
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
            Left            =   165
            TabIndex        =   20
            Top             =   555
            Width           =   1800
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "N° máximo lineas"
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
            Left            =   165
            TabIndex        =   19
            Top             =   945
            Width           =   1380
         End
      End
      Begin VB.Frame Frame 
         Height          =   3465
         Index           =   0
         Left            =   615
         TabIndex        =   9
         Top             =   3690
         Width           =   8460
         Begin VB.CommandButton cmdProcesaOtraVezHC 
            Caption         =   "Vuelve a procesar Historias con problemas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   150
            TabIndex        =   14
            Top             =   2730
            Width           =   7950
         End
         Begin VB.CheckBox chkSoundex 
            Caption         =   $"AHhcNOusadas.frx":12EE
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   180
            Picture         =   "AHhcNOusadas.frx":13E0
            TabIndex        =   13
            Top             =   1905
            Value           =   1  'Checked
            Width           =   8175
         End
         Begin VB.CheckBox chkPacientesIgualesSinDNI 
            Caption         =   $"AHhcNOusadas.frx":16F2
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   180
            Picture         =   "AHhcNOusadas.frx":1790
            TabIndex        =   12
            Top             =   1355
            Value           =   1  'Checked
            Width           =   8070
         End
         Begin VB.CheckBox chkDNIiguales 
            Caption         =   "DNIs iguales y Nros. De H/C diferentes teniendo datos de personas diferentes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            Picture         =   "AHhcNOusadas.frx":1AA2
            TabIndex        =   11
            Top             =   820
            Value           =   1  'Checked
            Width           =   8085
         End
         Begin VB.CheckBox chkDNIyPacienteIgual 
            Caption         =   $"AHhcNOusadas.frx":1DB4
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   180
            Picture         =   "AHhcNOusadas.frx":1E5C
            TabIndex        =   10
            Top             =   255
            Value           =   1  'Checked
            Width           =   8115
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   195
            Left            =   210
            TabIndex        =   22
            Top             =   3225
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin Threed.SSOption optHCtemporal 
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   2190
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   661
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Historias Temporales (para crearle Historia Clínica definitiva)"
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
         Enabled         =   0   'False
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
         Left            =   135
         Picture         =   "AHhcNOusadas.frx":216E
         TabIndex        =   4
         Top             =   7170
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin Threed.SSOption hcNoUsadas 
         Height          =   345
         Left            =   300
         TabIndex        =   6
         Top             =   150
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   609
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Historias Clínicas NO usadas (vacias)"
      End
      Begin Threed.SSOption optHCproblemas 
         Height          =   375
         Left            =   300
         TabIndex        =   7
         Top             =   2950
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   661
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Historias con problemas en la migración de Pacientes hacia SisGalenPlus (nª repetidos, pac repetidos)"
      End
      Begin Threed.SSOption optHCconProblemas 
         Height          =   375
         Left            =   300
         TabIndex        =   8
         Top             =   3330
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   661
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Historias con Problemas"
         Value           =   -1
      End
      Begin Threed.SSOption optHistoriasNuevas 
         Height          =   375
         Left            =   300
         TabIndex        =   25
         Top             =   2570
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   661
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de Historias que cambiaron por DNI"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   1
      Top             =   7680
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AHhcNOusadas.frx":2480
         DownPicture     =   "AHhcNOusadas.frx":28E0
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3210
         Picture         =   "AHhcNOusadas.frx":2D55
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AHhcNOusadas.frx":31CA
         DownPicture     =   "AHhcNOusadas.frx":368E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4740
         Picture         =   "AHhcNOusadas.frx":3B7A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "AHhcNOusadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Historias NO USADAS
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New RptAHhcNOusadas
        If hcNoUsadas.Value = True Then
           oRpt.CreaDatosParaReporte IIf(chkExcel.Value = 1, True, False), hcNoUsadas.Caption, ml_TextoDelFiltro, Val(txtHcMaxLinea.Text), Val(txtNlineas.Text), IIf(chkOrden.Value = 1, True, False), Me.hwnd
        ElseIf optHCtemporal.Value = True Then
           oRpt.CreaDatosParaReporteHcTemporales IIf(chkExcel.Value = 1, True, False), Me.optHCtemporal.Caption, ml_TextoDelFiltro, Me.hwnd
        ElseIf Me.optHCconProblemas.Value = True Or optHCconProblemas.Value = True Then
           CreaDatosParaReporteHcConProblemas True, "Lista de Historias con Problemas", "", Me.hwnd
        ElseIf optHistoriasNuevas.Value = True Then
           CrearDatosHistoriasIgualDNI
        End If
        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub

Sub CrearDatosHistoriasIgualDNI()
     Dim mrs_Tmp As New Recordset
     Dim lcSql As String, lcPie As String
     Set mrs_Tmp = mo_ReglasArchivoClinico.HistoriasClinicasSegunFiltro(" where not " & _
                   " (historiasclinicas.fechaUltimoCambioHistoria is null) order by historiasclinicas.fechaUltimoCambioHistoria desc")
     lcPie = Trim(Str(mrs_Tmp.RecordCount))
     If Val(lcPie) = 0 Then
        MsgBox "No existen HC", vbInformation, "Reporte"
     Else
        lcPie = "N° Registros: " & lcPie
        If chkExcel.Value = 1 Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, optHistoriasNuevas.Caption, "", "", Me.hwnd
           Set mo_ReglasReportes = Nothing
        End If
     End If
     Set mrs_Tmp = Nothing

End Sub
Sub CreaDatosParaReporteHcConProblemas(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, lnHwnd As Long)
     Dim mrs_Tmp As New Recordset
     Dim lcSql As String, lcPie As String
     Set mrs_Tmp = PacientesConHistoriasConProblemas
     lcPie = Trim(Str(mrs_Tmp.RecordCount))
     If Val(lcPie) = 0 Then
        MsgBox "No existen HC", vbInformation, "Reporte"
     Else
        lcPie = "N° Registros: " & lcPie
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, "", lnHwnd
           Set mo_ReglasReportes = Nothing
        End If
     End If
     Set mrs_Tmp = Nothing
     
End Sub

Function PacientesConHistoriasConProblemas() As Recordset
'   Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "lolcliProblemasHCSeleccionarTodos"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    
    Set PacientesConHistoriasConProblemas = oRsTmp1
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = ""
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



Private Sub cmdProcesaOtraVezHC_Click()
    On Error GoTo ProcHC
    If MsgBox("Está seguro ?", vbYesNo, "") = vbYes Then
       Me.MousePointer = 11
       cmdProcesaOtraVezHC.Enabled = False
       Dim oRsLolcli As New Recordset
       Dim oRsReporte5 As New Recordset
       Dim oRsPacientes As New Recordset
       Dim oConexion As New Connection
       Dim lnKEY As Long, lcError As String, lnMaximoRegistros As Long
       Dim lcSql As String, lbProcesa As Boolean
       Dim lnNroHistorias As Integer, lcIdPacienteOkey As String
       Dim AP1 As String, AM1 As String, N1 As String
       Dim AP2 As String, AM2 As String, N2 As String, lnIdPaciente As Long
       Dim NerrorAP As Integer, NerrorAM As Integer, NerrorN As Integer, lnFor As Integer
       oConexion.CommandTimeout = 900
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       
       If chkDNIyPacienteIgual.Value = 1 Then
            mo_ReglasArchivoClinico.AutogeneradoConNULL oConexion
       End If
       If chkDelDNIdefault.Value = 1 And Len(txtDNIusadoDefault.Text) = 8 Then
          lcSql = "update pacientes set nroDocumento=null,idTipoNumeracion=null where nroDocumento='" & _
                 Me.txtDNIusadoDefault.Text & "'"
          oRsLolcli.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       End If
       
       lcSql = "delete from lolcliProblemasHC where idPaciente>0"
       oRsLolcli.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       lcSql = "select * from lolcliProblemasHC"
       oRsLolcli.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
       mo_ReglasArchivoClinico.CreaTemporalPacientesAdepurar
       lcSql = "select * from reporte5"
       oRsReporte5.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'oRsReporte5.Filter = "idpaciente=3687"
       lnMaximoRegistros = oRsReporte5.RecordCount
       lnKEY = 1
       ProgressBar2.Min = 0
       ProgressBar2.Max = lnMaximoRegistros + 3
       ProgressBar2.Value = 0
       
       
       oRsReporte5.MoveFirst
       Do While Not oRsReporte5.EOF

'If ProgressBar2.Value = 18000 Then
'Exit Do
'End If
'Me.txtHcMaxLinea.Text = ProgressBar2.Value

          DoEvents: ProgressBar2.Value = ProgressBar2.Value + 1: Me.Refresh
          If BuscaEnSiSeProcesoHistoria(oRsLolcli, oRsReporte5!IdPaciente) = True Then
                lbProcesa = False
                If lbProcesa = False And chkDNIyPacienteIgual.Value = 1 Then
                    If Not IsNull(oRsReporte5!NroDocumento) And Not IsNull(oRsReporte5!idDocIdentidad) And Not IsNull(oRsReporte5!FechaNacimiento) Then
                        If oRsReporte5!NroDocumento <> "" Then
                           lcSql = "select * from Pacientes where IdTipoNumeracion <=2 and not (NroHistoriaClinica is null) and " & _
                                   "  apellidoPaterno='" & oRsReporte5!ApellidoPaterno & "' and apellidoMaterno='" & oRsReporte5!apellidoMaterno & _
                                   "' and PrimerNombre='" & oRsReporte5!PrimerNombre & _
                                   "' and convert(char(10),FechaNacimiento,103)='" & Format(oRsReporte5!FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                   "' and idTipoSexo=" & oRsReporte5!idTipoSexo & " and nroDocumento='" & oRsReporte5!NroDocumento & _
                                   "' and idDocIdentidad=" & oRsReporte5!idDocIdentidad
                           oRsPacientes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                           If oRsPacientes.RecordCount > 1 Then
                              lcError = "DNIigualesDatosDelPacienteIguales"
                              oRsPacientes.MoveFirst
                              Do While Not oRsPacientes.EOF
                                 oRsLolcli.AddNew
                                 oRsLolcli!pacHis = Trim(Str(oRsPacientes!NroHistoriaClinica))
                                 oRsLolcli!pacPat = Left(oRsPacientes!ApellidoPaterno, 30)
                                 oRsLolcli!pacMat = Left(oRsPacientes!apellidoMaterno, 30)
                                 oRsLolcli!pacNam = Left(oRsPacientes!PrimerNombre + IIf(IsNull(oRsPacientes!SegundoNombre), "", " " & _
                                                                                         oRsPacientes!SegundoNombre), 50)
                                 oRsLolcli!nroHistoriaGalenhos = Trim(Str(oRsPacientes!NroHistoriaClinica))
                                 oRsLolcli!dni = oRsPacientes!NroDocumento
                                 oRsLolcli!IdPaciente = Trim(Str(lnKEY))
                                 oRsLolcli!FechaNacimiento = IIf(IsNull(oRsPacientes!FechaNacimiento), 0, oRsPacientes!FechaNacimiento)
                                 oRsLolcli!idTipoSexo = IIf(IsNull(oRsPacientes!idTipoSexo), 0, oRsPacientes!idTipoSexo)
                                 oRsLolcli!autogeneradoGalenHos = Trim(Str(oRsPacientes!IdPaciente)) & "/" & lcError
                                 oRsLolcli.Update
                                 oRsPacientes.MoveNext
                              Loop
                              AgregaLineaVaciaEnLolcli oRsLolcli, lnKEY, lcError
                              lnKEY = lnKEY + 1
                              lbProcesa = True
                           End If
                           oRsPacientes.Close
                       End If
                    End If
                End If
                If lbProcesa = False And chkDNIiguales.Value = 1 Then
                    If Not IsNull(oRsReporte5!NroDocumento) And Not IsNull(oRsReporte5!idDocIdentidad) Then
                        If oRsReporte5!NroDocumento <> "" Then
                           lcSql = "select * from Pacientes where IdTipoNumeracion <=2 and not (NroHistoriaClinica is null) and " & _
                                   " nroDocumento='" & oRsReporte5!NroDocumento & "' and idDocIdentidad=" & oRsReporte5!idDocIdentidad
                           oRsPacientes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                           If oRsPacientes.RecordCount > 1 Then
                              lcError = "DNIigualesDatosDelPacienteDiferentes"
                              oRsPacientes.MoveFirst
                              Do While Not oRsPacientes.EOF
                                 oRsLolcli.AddNew
                                 oRsLolcli!pacHis = Trim(Str(oRsPacientes!NroHistoriaClinica))
                                 oRsLolcli!pacPat = Left(oRsPacientes!ApellidoPaterno, 30)
                                 oRsLolcli!pacMat = Left(oRsPacientes!apellidoMaterno, 30)
                                 oRsLolcli!pacNam = Left(oRsPacientes!PrimerNombre + IIf(IsNull(oRsPacientes!SegundoNombre), "", " " & _
                                                                                         oRsPacientes!SegundoNombre), 50)
                                 oRsLolcli!nroHistoriaGalenhos = Trim(Str(oRsPacientes!NroHistoriaClinica))
                                 oRsLolcli!dni = oRsPacientes!NroDocumento
                                 oRsLolcli!IdPaciente = Trim(Str(lnKEY))
                                 oRsLolcli!FechaNacimiento = IIf(IsNull(oRsPacientes!FechaNacimiento), 0, oRsPacientes!FechaNacimiento)
                                 oRsLolcli!idTipoSexo = IIf(IsNull(oRsPacientes!idTipoSexo), 0, oRsPacientes!idTipoSexo)
                                 oRsLolcli!autogeneradoGalenHos = Trim(Str(oRsPacientes!IdPaciente)) & "/" & lcError
                                 oRsLolcli.Update
                                 oRsPacientes.MoveNext
                              Loop
                              AgregaLineaVaciaEnLolcli oRsLolcli, lnKEY, lcError
                              lnKEY = lnKEY + 1
                              lbProcesa = True
                           End If
                           oRsPacientes.Close
                       End If
                    End If
                End If
                If lbProcesa = False And chkPacientesIgualesSinDNI.Value = 1 And Not IsNull(oRsReporte5!FechaNacimiento) And (IsNull(oRsReporte5!NroDocumento) Or oRsReporte5!NroDocumento = "") Then
                      lcSql = "select * from Pacientes where IdTipoNumeracion <=2 and (idDocIdentidad is null or idDocIdentidad=10) and " & _
                              "  apellidoPaterno='" & oRsReporte5!ApellidoPaterno & "' and apellidoMaterno='" & oRsReporte5!apellidoMaterno & _
                              "' and PrimerNombre='" & oRsReporte5!PrimerNombre & _
                              "' and convert(char(10),FechaNacimiento,103)='" & Format(oRsReporte5!FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                              "' and idTipoSexo=" & oRsReporte5!idTipoSexo
                      oRsPacientes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                      If oRsPacientes.RecordCount > 1 Then
                         lcError = "DatosDelPacienteIgualesSinDNI"
                         oRsPacientes.MoveFirst
                         Do While Not oRsPacientes.EOF
                            oRsLolcli.AddNew
                            oRsLolcli!pacHis = Trim(Str(oRsPacientes!NroHistoriaClinica))
                            oRsLolcli!pacPat = Left(oRsPacientes!ApellidoPaterno, 30)
                            oRsLolcli!pacMat = Left(oRsPacientes!apellidoMaterno, 30)
                            oRsLolcli!pacNam = Left(oRsPacientes!PrimerNombre + IIf(IsNull(oRsPacientes!SegundoNombre), "", " " & _
                                                                                    oRsPacientes!SegundoNombre), 50)
                            oRsLolcli!nroHistoriaGalenhos = Trim(Str(oRsPacientes!NroHistoriaClinica))
                            oRsLolcli!dni = oRsPacientes!NroDocumento
                            oRsLolcli!IdPaciente = Trim(Str(lnKEY))
                            oRsLolcli!autogeneradoGalenHos = Trim(Str(oRsPacientes!IdPaciente)) & "/" & lcError
                            oRsLolcli!FechaNacimiento = IIf(IsNull(oRsPacientes!FechaNacimiento), 0, oRsPacientes!FechaNacimiento)
                            oRsLolcli!idTipoSexo = IIf(IsNull(oRsPacientes!idTipoSexo), 0, oRsPacientes!idTipoSexo)
                            oRsLolcli.Update
                            oRsPacientes.MoveNext
                         Loop
                         AgregaLineaVaciaEnLolcli oRsLolcli, lnKEY, lcError
                         lnKEY = lnKEY + 1
                         lbProcesa = True
                      End If
                      oRsPacientes.Close
                End If
                If lbProcesa = False And chkSoundex.Value = 1 And Not IsNull(oRsReporte5!autogenerado) Then
                      lcSql = "select * from Pacientes where IdTipoNumeracion <=2  and " & _
                              "  autogenerado='" & oRsReporte5!autogenerado & "'"
                      oRsPacientes.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                      If oRsPacientes.RecordCount > 1 Then
                         
                         lcError = "DatosDelPacienteCasiIguales"
                         lnNroHistorias = 0
                         lcIdPacienteOkey = "/"
                         AP1 = oRsReporte5!ApellidoPaterno
                         AM1 = oRsReporte5!apellidoMaterno
                         N1 = oRsReporte5!PrimerNombre
                         oRsPacientes.MoveFirst
                         Do While Not oRsPacientes.EOF
                            lnIdPaciente = oRsPacientes!IdPaciente
                            AP2 = oRsPacientes!ApellidoPaterno
                            AM2 = oRsPacientes!apellidoMaterno
                            N2 = oRsPacientes!PrimerNombre
                         
                            NerrorAP = 0
                            For lnFor = 1 To Len(AP1)
                                If Mid(AP1, lnFor, 1) <> Mid(AP2, lnFor, 1) Then
                                   NerrorAP = NerrorAP + 1
                                End If
                            Next
                            
                            NerrorAM = 0
                            For lnFor = 1 To Len(AM1)
                                If Mid(AM1, lnFor, 1) <> Mid(AM2, lnFor, 1) Then
                                   NerrorAM = NerrorAM + 1
                                End If
                            Next

                            NerrorN = 0
                            For lnFor = 1 To Len(N1)
                                If Mid(N1, lnFor, 1) <> Mid(N2, lnFor, 1) Then
                                   NerrorN = NerrorN + 1
                                End If
                            Next

                            If NerrorAP <= 3 And NerrorAM <= 3 And NerrorN <= 3 Then
                                lcIdPacienteOkey = lcIdPacienteOkey + Trim(Str(lnIdPaciente)) + "/"
                                lnNroHistorias = lnNroHistorias + 1
                            End If

                            oRsPacientes.MoveNext
                         Loop
                         If lnNroHistorias > 1 Then
                            oRsPacientes.MoveFirst
                            Do While Not oRsPacientes.EOF
                               If InStr(lcIdPacienteOkey, Trim(Str(oRsPacientes!IdPaciente))) > 0 Then
                                    oRsLolcli.AddNew
                                    oRsLolcli!pacHis = Trim(Str(oRsPacientes!NroHistoriaClinica))
                                    oRsLolcli!pacPat = Left(oRsPacientes!ApellidoPaterno, 30)
                                    oRsLolcli!pacMat = Left(oRsPacientes!apellidoMaterno, 30)
                                    oRsLolcli!pacNam = Left(oRsPacientes!PrimerNombre + IIf(IsNull(oRsPacientes!SegundoNombre), "", " " & _
                                                                                            oRsPacientes!SegundoNombre), 50)
                                    oRsLolcli!nroHistoriaGalenhos = Trim(Str(oRsPacientes!NroHistoriaClinica))
                                    oRsLolcli!dni = oRsPacientes!NroDocumento
                                    oRsLolcli!IdPaciente = Trim(Str(lnKEY))
                                    oRsLolcli!autogeneradoGalenHos = Trim(Str(oRsPacientes!IdPaciente)) & "/" & lcError
                                    oRsLolcli!FechaNacimiento = IIf(IsNull(oRsPacientes!FechaNacimiento), 0, oRsPacientes!FechaNacimiento)
                                    oRsLolcli!idTipoSexo = IIf(IsNull(oRsPacientes!idTipoSexo), 0, oRsPacientes!idTipoSexo)
                                    oRsLolcli.Update
                               End If
                               oRsPacientes.MoveNext
                            Loop
                            AgregaLineaVaciaEnLolcli oRsLolcli, lnKEY, lcError
                            lnKEY = lnKEY + 1
                            lbProcesa = True
                         End If
                      End If
                      oRsPacientes.Close
                End If
          End If
          oRsReporte5.MoveNext
       Loop
       oRsLolcli.Close
       oRsReporte5.Close
       oConexion.Close
       optHCconProblemas.Value = True
       Set oRsLolcli = Nothing
       Set oRsReporte5 = Nothing
       Set oRsPacientes = Nothing
       Set oConexion = Nothing
       Me.MousePointer = 1
       btnAceptar_Click
    End If
    Exit Sub
ProcHC:
    MsgBox Err.Description
    Set oRsLolcli = Nothing
    Set oRsReporte5 = Nothing
    Set oRsPacientes = Nothing
    Set oConexion = Nothing
    Me.MousePointer = 1
    Exit Sub
    Resume
End Sub

Sub AgregaLineaVaciaEnLolcli(oRsLolcli As Recordset, lnKEY As Long, lcError)
                   oRsLolcli.AddNew
                   oRsLolcli!pacHis = "*****"
                   oRsLolcli!pacPat = "*****"
                   oRsLolcli!pacMat = "*****"
                   oRsLolcli!pacNam = "*****"
                   oRsLolcli!nroHistoriaGalenhos = "*****"
                   oRsLolcli!dni = "*****"
                   oRsLolcli!IdPaciente = Trim(Str(lnKEY))
                   oRsLolcli!autogeneradoGalenHos = "*****"
                   oRsLolcli.Update
End Sub

Function BuscaEnSiSeProcesoHistoria(oRsLolcli As Recordset, lnIdPaciente As Long) As Boolean
    BuscaEnSiSeProcesoHistoria = True
    If oRsLolcli.RecordCount > 0 Then
       oRsLolcli.MoveFirst
       Do While Not oRsLolcli.EOF
          If oRsLolcli!autogeneradoGalenHos <> "*****" And InStr(oRsLolcli!autogeneradoGalenHos, "/") > 0 Then
                If Val(Left(oRsLolcli!autogeneradoGalenHos, InStr(oRsLolcli!autogeneradoGalenHos, "/") - 1)) = lnIdPaciente Then
                   BuscaEnSiSeProcesoHistoria = False
                   Exit Do
                End If
          End If
          oRsLolcli.MoveNext
       Loop
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_Teclado = Nothing
    Set mo_Formulario = Nothing
End Sub



