VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Begin VB.UserControl UcEpisodioClinico 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   ScaleHeight     =   945
   ScaleWidth      =   4800
   Begin VB.Frame FraEpisodio 
      Caption         =   "Episodio Clínico "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.CheckBox chkEpisodioCerrar 
         Alignment       =   1  'Right Justify
         Caption         =   "Cierre de Episodio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2880
         TabIndex        =   5
         Top             =   570
         Width           =   1755
      End
      Begin VB.CheckBox chkEpisodioNew 
         Alignment       =   1  'Right Justify
         Caption         =   "Nuevo Episodio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1755
      End
      Begin PVCOMBOLibCtl.PVComboBox cmdEpisodiosHistoricos 
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   825
         _Version        =   524288
         _cx             =   1455
         _cy             =   582
         Appearance      =   1
         Enabled         =   -1  'True
         BackColor       =   16777215
         ForeColor       =   0
         Locked          =   0   'False
         Style           =   0
         Sorted          =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowPictures    =   0   'False
         ColumnHeaders   =   -1  'True
         PrimaryColumn   =   1
         VisibleItems    =   10
         ColumnHeaderHeight=   20
         ListMember      =   ""
         ColumnHeaderForeColor=   0
         ColumnHeaderBackColor=   13160660
         SelectedForeColor=   16777215
         SelectedBackColor=   6956042
         AlternateBackColor=   16777215
         ItemLabelStyle  =   1
         ItemLabelType   =   0
         ItemLabelWidth  =   40
         ItemLabelForeColor=   0
         ItemLabelBackColor=   13160660
         ColumnHeaderStyle=   1
         VerticalGridLines=   -1  'True
         HorizontalGridLines=   -1  'True
         ColumnResize    =   0   'False
         ItemLabelResize =   0   'False
         AllowDBAutoConfig=   0   'False
         GridLineColor   =   13421772
         List            =   ""
         NullString      =   "[NULL]"
         DropShadow      =   -1  'True
         Text            =   ""
         SortOnColumnHeaderClick=   0   'False
         DropEffect      =   1
         ColumnCount     =   8
         Column0.Heading =   "Id"
         Column0.Width   =   10
         Column0.Alignment=   0
         Column0.Hidden  =   0   'False
         Column0.Name    =   "NroEpisodio"
         Column0.Format  =   ""
         Column0.Bound   =   -1  'True
         Column0.Locked  =   0   'False
         Column0.HeaderAlignment=   0
         Column1.Heading =   "F.Apertura"
         Column1.Width   =   38
         Column1.Alignment=   0
         Column1.Hidden  =   0   'False
         Column1.Name    =   "FechaApertura"
         Column1.Format  =   ""
         Column1.Bound   =   -1  'True
         Column1.Locked  =   0   'False
         Column1.HeaderAlignment=   0
         Column2.Heading =   "F. Cierre"
         Column2.Width   =   38
         Column2.Alignment=   0
         Column2.Hidden  =   0   'False
         Column2.Name    =   "FechaCierre"
         Column2.Format  =   ""
         Column2.Bound   =   -1  'True
         Column2.Locked  =   0   'False
         Column2.HeaderAlignment=   0
         Column3.Heading =   "Diagnóstico"
         Column3.Width   =   200
         Column3.Alignment=   0
         Column3.Hidden  =   0   'False
         Column3.Name    =   "Dx"
         Column3.Format  =   ""
         Column3.Bound   =   -1  'True
         Column3.Locked  =   0   'False
         Column3.HeaderAlignment=   0
         Column4.Heading =   "Tipo Servicio"
         Column4.Width   =   20
         Column4.Alignment=   0
         Column4.Hidden  =   0   'False
         Column4.Name    =   "TipoServicio"
         Column4.Format  =   ""
         Column4.Bound   =   -1  'True
         Column4.Locked  =   0   'False
         Column4.HeaderAlignment=   0
         Column5.Heading =   "Servicio/Consultorio"
         Column5.Width   =   60
         Column5.Alignment=   0
         Column5.Hidden  =   0   'False
         Column5.Name    =   "Servicio"
         Column5.Format  =   ""
         Column5.Bound   =   -1  'True
         Column5.Locked  =   0   'False
         Column5.HeaderAlignment=   0
         Column6.Heading =   "F. Cuenta"
         Column6.Width   =   38
         Column6.Alignment=   0
         Column6.Hidden  =   0   'False
         Column6.Name    =   "CuentaFecha"
         Column6.Format  =   ""
         Column6.Bound   =   -1  'True
         Column6.Locked  =   0   'False
         Column6.HeaderAlignment=   0
         Column7.Heading =   "N° Cuenta"
         Column7.Width   =   40
         Column7.Alignment=   0
         Column7.Hidden  =   0   'False
         Column7.Name    =   "CuentaNro"
         Column7.Format  =   ""
         Column7.Bound   =   -1  'True
         Column7.Locked  =   0   'False
         Column7.HeaderAlignment=   0
         SortKey1.Column =   -1
         SortKey1.Ascending=   -1  'True
         SortKey1.CaseInsensitive=   -1  'True
         SortKey2.Column =   -1
         SortKey2.Ascending=   -1  'True
         SortKey2.CaseInsensitive=   -1  'True
         SortKey3.Column =   -1
         SortKey3.Ascending=   -1  'True
         SortKey3.CaseInsensitive=   -1  'True
         BoundColumn     =   ""
         Border          =   -1  'True
         VertAlign       =   1
         Format          =   ""
      End
      Begin VB.Label lblEpisodioElegido 
         AutoSize        =   -1  'True
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2580
         TabIndex        =   2
         Top             =   270
         Width           =   135
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Episodio histórico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   1395
      End
   End
End
Attribute VB_Name = "UcEpisodioClinico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para elegir episodio clínico
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Dim oRsHistoricos As New Recordset

Dim ml_IdPaciente As Long
Dim ml_idAtencion As Long
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim lbTieneTodosEpisodiosCerrados As Boolean

Property Let idPaciente(lValue As Long)
    ml_IdPaciente = lValue
End Property
Property Get idPaciente() As Long
    idPaciente = ml_IdPaciente
End Property

Property Let idAtencion(lValue As Long)
    ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
    idAtencion = ml_idAtencion
End Property

Property Get idEpisodio() As Long
    idEpisodio = Val(cmdEpisodiosHistoricos.BoundText)
End Property
Property Get lbCierreEpisodio() As Boolean
    lbCierreEpisodio = IIf(chkEpisodioCerrar.Value = 1, True, False)
End Property

Property Get lbNuevoEpisodio() As Boolean
    lbNuevoEpisodio = IIf(chkEpisodioNew.Value = 1, True, False)
End Property



Private Sub chkEpisodioCerrar_Click()
    If chkEpisodioNew.Value = 1 And chkEpisodioCerrar.Value = 1 Then
       cmdEpisodiosHistoricos.Text = ""
       lblEpisodioElegido.Caption = ""
    End If
End Sub

Private Sub chkEpisodioNew_Click()
    If chkEpisodioNew.Value = 1 Then
       cmdEpisodiosHistoricos.Text = ""
       lblEpisodioElegido.Caption = ""
    End If
End Sub

Private Sub cmdEpisodiosHistoricos_Click()
    EpisodioElegido
    SendKeys "{tab}"
    
End Sub

Private Sub cmdEpisodiosHistoricos_DblClick()
    EpisodioElegido
    SendKeys "{tab}"
End Sub

Private Sub cmdEpisodiosHistoricos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       EpisodioElegido
    End If
End Sub


Sub EpisodioElegido()
    Dim oCampos() As String
    On Error Resume Next
    oCampos = Split(cmdEpisodiosHistoricos.List(cmdEpisodiosHistoricos.ListIndex), "|")
    If UCase(Trim(oCampos(2))) = "[NULL]" Then
        lblEpisodioElegido.Caption = oCampos(3)
        cmdEpisodiosHistoricos.Text = oCampos(0)
        If UserControl.chkEpisodioNew.Value = 1 Then
           UserControl.chkEpisodioNew.Value = 0
        End If
    Else
        lblEpisodioElegido.Caption = ""
        cmdEpisodiosHistoricos.Text = oCampos(0)
        MsgBox "Solo podrá elegir EPISODIOS sin FECHA CIERRE", vbInformation, "Epidosio Clínico"
    End If
End Sub

Private Sub cmdEpisodiosHistoricos_LostFocus()
    EpisodioElegido
End Sub

Private Sub UserControl_Resize()
    UserControl.FraEpisodio.Width = UserControl.Width - 20
    UserControl.FraEpisodio.Height = UserControl.Height - 20
End Sub

Sub CargaEpisodiosHistoricos()
    'llena temporal con Historicos de atenciones
    Dim oRsTmp1 As New Recordset
    Dim oDODiagnostico As New DODiagnostico
    Dim oConexion As New Connection
    lbTieneTodosEpisodiosCerrados = True
    If ml_IdPaciente > 0 Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        If oRsHistoricos.State = 1 Then Set oRsHistoricos = Nothing
        With oRsHistoricos
            .Fields.Append "NroEpisodio", adInteger
            .Fields.Append "FechaApertura", adDate, , adFldIsNullable
            .Fields.Append "FechaCierre", adDate, , adFldIsNullable
            .Fields.Append "Dx", adVarChar, 300, adFldIsNullable
            .Fields.Append "TipoServicio", adVarChar, 100, adFldIsNullable
            .Fields.Append "Servicio", adVarChar, 200, adFldIsNullable
            .Fields.Append "CuentaFecha", adDate, , adFldIsNullable
            .Fields.Append "CuentaNro", adInteger
            .LockType = adLockOptimistic
            .Open
        End With
        Set oRsTmp1 = mo_ReglasAdmision.AtencionesEpisodiosDetalleSeleccionarXpaciente(ml_IdPaciente, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                Set oDODiagnostico = mo_AdminFacturacion.DevuelveDxAltaMedica(oRsTmp1!idAtencion, oRsTmp1!idTipoServicio, oConexion)
                oRsHistoricos.AddNew
                oRsHistoricos.Fields!NroEpisodio = oRsTmp1!idEpisodio
                oRsHistoricos.Fields!FechaApertura = oRsTmp1!FechaApertura
                oRsHistoricos.Fields!fechaCierre = oRsTmp1!fechaCierre
                oRsHistoricos.Fields!dx = Trim(oDODiagnostico.CodigoCIE2004) & " " & oDODiagnostico.Descripcion
                oRsHistoricos.Fields!TipoServicio = oRsTmp1!TipoServicio
                If oRsTmp1.Fields!idTipoServicio = sghConsultaExterna Then
                   oRsHistoricos.Fields!Servicio = oRsTmp1!ServIng
                Else
                   oRsHistoricos.Fields!Servicio = oRsTmp1!Servicio
                End If
                oRsHistoricos.Fields!CuentaFecha = oRsTmp1!FechaIngreso
                oRsHistoricos.Fields!CuentaNro = oRsTmp1!idCuentaAtencion
                oRsHistoricos.Update
                If IsNull(oRsTmp1!fechaCierre) Then
                   lbTieneTodosEpisodiosCerrados = False
                End If
                oRsTmp1.MoveNext
           Loop
        End If
        oConexion.Close
        Set cmdEpisodiosHistoricos.ListSource = oRsHistoricos
    Else
        lbTieneTodosEpisodiosCerrados = True
    End If
    Set oRsTmp1 = Nothing
    Set oDODiagnostico = Nothing
    Set oConexion = Nothing
    'If lbTieneTodosEpisodiosCerrados = True Then
    '   chkEpisodioNew.Value = 1
    '   chkEpisodioCerrar.Value = 0
       
    'End If
End Sub

Sub Inicializar()
    Dim oRsTmp1 As New Recordset
    Dim oDODiagnostico As New DODiagnostico
    Dim oConexion As New Connection
    lbTieneTodosEpisodiosCerrados = False
    chkEpisodioNew.Value = 0
    chkEpisodioNew.Enabled = True
    chkEpisodioCerrar.Value = 0
    chkEpisodioCerrar.Enabled = True
    Set cmdEpisodiosHistoricos.ListSource = Nothing
    cmdEpisodiosHistoricos.Text = ""
    lblEpisodioElegido.Caption = ""
End Sub

Function ValidaReglas(lcFicha As String) As Boolean
    ValidaReglas = False
    Dim lcMensaje As String
    lcMensaje = ""
    If chkEpisodioCerrar.Value = 0 And chkEpisodioNew.Value = 0 And lblEpisodioElegido.Caption = "" Then
       lcMensaje = lcMensaje & "Falta elejir datos para el EPISODIO CLINICO " & lcFicha & Chr(13)
    End If
    If chkEpisodioCerrar.Value = 1 And chkEpisodioNew.Value = 0 And lblEpisodioElegido.Caption = "" Then
       lcMensaje = lcMensaje & "Falta elejir el EPISODIO HISTORICO para poder CERRAR el EPISODIO CLINICO " & lcFicha & Chr(13)
    End If
    If lcMensaje = "" Then
       ValidaReglas = True
    Else
       MsgBox lcMensaje, vbInformation, "Episodio Clínico"
    End If
End Function

Public Function DevuelveDatosElegidos() As EpisodioClinico
    Dim oEpisodioClinico As EpisodioClinico
    oEpisodioClinico.idEpisodio = Val(cmdEpisodiosHistoricos.Text)
    oEpisodioClinico.lbCierreEpisodio = IIf(chkEpisodioCerrar.Value = 1, True, False)
    oEpisodioClinico.lbNuevoEpisodio = IIf(chkEpisodioNew.Value = 1, True, False)
    DevuelveDatosElegidos = oEpisodioClinico
End Function

Sub CargarDatosAlosControles(oConexion As Connection)
    If ml_idAtencion > 0 And ml_IdPaciente > 0 Then
       Dim oRsTmp1 As New Recordset
       Set oRsTmp1 = mo_ReglasAdmision.AtencionesEpisodiosDetalleSeleccionarXpaciente(ml_IdPaciente, oConexion)
       oRsTmp1.Filter = "idAtencion=" & ml_idAtencion
       If oRsTmp1.RecordCount > 0 Then
          If IsNull(oRsTmp1!fechaCierre) Then
             chkEpisodioNew.Value = 1
          Else
             chkEpisodioCerrar.Value = 1
             If oRsTmp1.RecordCount = 1 Then
                chkEpisodioNew.Value = 1
             End If
          End If
          cmdEpisodiosHistoricos.BoundText = oRsTmp1!idEpisodio
          If oRsHistoricos.RecordCount > 0 Then
             oRsHistoricos.MoveFirst
             oRsHistoricos.Find "NroEpisodio=" & oRsTmp1!idEpisodio
             If Not oRsHistoricos.EOF Then
                lblEpisodioElegido.Caption = oRsHistoricos!dx
             End If
          End If
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
    End If
End Sub

Sub Limpiar()
   cmdEpisodiosHistoricos.Text = ""
   lblEpisodioElegido.Caption = ""
   chkEpisodioNew.Value = 0
   chkEpisodioCerrar.Value = 0
End Sub
