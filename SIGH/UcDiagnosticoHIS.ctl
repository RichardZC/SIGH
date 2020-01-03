VERSION 5.00
Object = "{0FAA9261-2AF4-11D3-9995-00A0CC3A27A9}#1.0#0"; "PVCombo.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl UcDiagnosticoHIS 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   ScaleHeight     =   5070
   ScaleWidth      =   10485
   Begin VB.Frame fraDiagnostico 
      Caption         =   "Diagnósticos     ( F1=Todos Dx )"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   0
      TabIndex        =   4
      Top             =   15
      Width           =   10455
      Begin VB.CommandButton btnBusquedaDiagnostico 
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
         Left            =   2535
         Picture         =   "UcDiagnosticoHIS.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Buscar"
         Top             =   225
         Width           =   375
      End
      Begin VB.ComboBox cmbConsultorios 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1005
         Width           =   3780
      End
      Begin VB.Frame fraLab 
         Caption         =   "Lab (HIS)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   8460
         TabIndex        =   9
         Top             =   150
         Width           =   1920
         Begin PVCOMBOLibCtl.PVComboBox cmbLabHis 
            Height          =   330
            Left            =   1200
            TabIndex        =   11
            Top             =   180
            Width           =   690
            _Version        =   524288
            _cx             =   1217
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
               Size            =   8.25
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
            ItemLabelWidth  =   20
            ItemLabelForeColor=   0
            ItemLabelBackColor=   13160660
            ColumnHeaderStyle=   0
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
            ColumnCount     =   3
            Column0.Heading =   "Id"
            Column0.Width   =   10
            Column0.Alignment=   0
            Column0.Hidden  =   -1  'True
            Column0.Name    =   "IdHisSituacio"
            Column0.Format  =   ""
            Column0.Bound   =   -1  'True
            Column0.Locked  =   0   'False
            Column0.HeaderAlignment=   0
            Column1.Heading =   "Valores"
            Column1.Width   =   35
            Column1.Alignment=   0
            Column1.Hidden  =   0   'False
            Column1.Name    =   "valores"
            Column1.Format  =   ""
            Column1.Bound   =   -1  'True
            Column1.Locked  =   0   'False
            Column1.HeaderAlignment=   0
            Column2.Heading =   "Descripción"
            Column2.Width   =   100
            Column2.Alignment=   0
            Column2.Hidden  =   0   'False
            Column2.Name    =   "descripcio"
            Column2.Format  =   ""
            Column2.Bound   =   -1  'True
            Column2.Locked  =   0   'False
            Column2.HeaderAlignment=   0
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
         Begin UltraGrid.SSUltraGrid grdLabHIS 
            Height          =   1050
            Left            =   75
            TabIndex        =   10
            Top             =   180
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   1852
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "grdLabHIS"
         End
      End
      Begin VB.TextBox txtIdDiagnostico 
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
         Left            =   1500
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox lblDescripcionDx 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2925
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   5505
      End
      Begin VB.ComboBox cmbIdTipoDiagnostico 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "UcDiagnosticoHIS.ctx":058A
         Left            =   1500
         List            =   "UcDiagnosticoHIS.ctx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   615
         Width           =   3780
      End
      Begin Threed.SSCommand btnAgregarDx 
         Height          =   465
         Left            =   5910
         TabIndex        =   3
         Top             =   930
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "UcDiagnosticoHIS.ctx":058E
         Caption         =   "Agregar"
         PictureAlignment=   9
      End
      Begin Threed.SSCommand btnQuitarDx 
         Height          =   465
         Left            =   7185
         TabIndex        =   6
         Top             =   930
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   820
         _Version        =   262144
         PictureFrames   =   1
         Picture         =   "UcDiagnosticoHIS.ctx":351A
         Caption         =   "Quitar"
         PictureAlignment=   9
         ShapeSize       =   1
      End
      Begin UltraGrid.SSUltraGrid grdDiagnosticos 
         Height          =   3465
         Left            =   105
         TabIndex        =   14
         Top             =   1470
         Width           =   10290
         _ExtentX        =   18150
         _ExtentY        =   6112
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de diagnósticos"
      End
      Begin VB.Label lblConsultorioActual 
         Caption         =   "......."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   975
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblConsultorio 
         AutoSize        =   -1  'True
         Caption         =   "UPS"
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
         TabIndex        =   12
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label lblDiagnostico 
         AutoSize        =   -1  'True
         Caption         =   "Tipo diagnóstico"
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
         TabIndex        =   8
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Diagnóstico"
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
         TabIndex        =   7
         Top             =   300
         Width           =   930
      End
   End
End
Attribute VB_Name = "UcDiagnosticoHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar Diagnósticos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idAtencion As Long
Dim ml_idUsuario As Long
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ms_MensajeError As String
Dim mrs_Diagnosticos As New ADODB.Recordset
Dim mrs_Labs As New Recordset
Dim ml_TipoDiagnostico As sghTiposDiagnostico
Dim mo_cmbIdTipoDiagnostico As New sighentidades.ListaDespleglable
Dim mo_cmbConsultorios As New sighentidades.ListaDespleglable
Dim ml_SexoPaciente As Long
Dim ml_EdadPaciente As Long
Dim ml_IdListBarItem As Long
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim ml_AScorrelativo As Long
Dim ml_idCuentaAtencion_actual As Long
Dim ml_Consultorio As String
Dim rsConsultorios As New Recordset
Dim mRs_Fua As New Recordset
Dim ml_IdFuenteFinanciamiento As Long
Dim rsFUAequiv As New Recordset
Dim lnFUAequiv As Long
Dim ml_PesoKg As Double

Dim ml_FechaNacimiento As Date
Dim ml_FechaAtencion As Date
Dim ml_ups As String
Dim ml_IdServicio As Long
Dim oRsTipoDx As New Recordset

Property Set oRsItemsElegidos(oValue As Recordset)
    Dim ml_oRsItemsElegidos As Recordset
    Dim oRsTmp1 As New Recordset
    Set ml_oRsItemsElegidos = oValue
    If ml_oRsItemsElegidos.RecordCount > 0 Then
       ml_oRsItemsElegidos.MoveFirst
       Do While Not ml_oRsItemsElegidos.EOF
                oRsTipoDx.MoveFirst
                oRsTipoDx.Find "IdSubclasificacionDx=" & ml_oRsItemsElegidos!elijaTipo
                Set oRsTmp1 = mo_AdminServiciosComunes.DiagnosticosSeleccionarXCodigo(Trim(ml_oRsItemsElegidos!Id))
                With mrs_Diagnosticos
                    .AddNew
                    .Fields!idTipoDiagnostico = ml_oRsItemsElegidos!elijaTipo
                    .Fields!DescripcionTipoDx = oRsTipoDx!DescripcionLarga
                    .Fields!idDiagnostico = oRsTmp1!idDiagnostico
                    .Fields!CodigoCIE2004 = ml_oRsItemsElegidos!Id
                    .Fields!Descripcion = ml_oRsItemsElegidos!nombre
                    .Fields!labConfHIS = ml_oRsItemsElegidos!ElijaLab
                    If ml_AScorrelativo = 0 Then
                        .Fields!Consultorio = ml_Consultorio
                        .Fields!idCuentaAtencion = ml_idCuentaAtencion_actual
                        .Fields!IdServicio = ml_IdServicio
                    Else
                        .Fields!Consultorio = ml_oRsItemsElegidos!Consultorio
                        .Fields!idCuentaAtencion = ml_oRsItemsElegidos!idCuentaAtencion
                        .Fields!IdServicio = ml_oRsItemsElegidos!IdServicio
                    End If
                    .Fields!FUA = ml_oRsItemsElegidos!FUA
                    .Fields!Grupo = ml_oRsItemsElegidos!Grupo
                    .Fields!SubGrupo = ml_oRsItemsElegidos!SubGrupo
                    .Update
                End With
                ml_oRsItemsElegidos.MoveNext
    
       Loop
     End If
     Set ml_oRsItemsElegidos = Nothing
     Set oRsTmp1 = Nothing
End Property

Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Let UPS(oValue As String)
    ml_ups = oValue
End Property

Property Let FechaAtencion(oValue As Date)
    ml_FechaAtencion = oValue
End Property
Property Let FechaNacimiento(oValue As Date)
    ml_FechaNacimiento = oValue
    
End Property

Property Let PesoKg(lValue As Double)
   ml_PesoKg = lValue
End Property
Property Let IdFuenteFinanciamiento(lValue As Long)
   ml_IdFuenteFinanciamiento = lValue
End Property



Property Set RsServiciosAtenSimultaneaFuaXcorrelativo(oValue As Recordset)
    Set mRs_Fua = oValue
    If mrs_Diagnosticos.RecordCount > 0 Then
        mRs_Fua.Filter = "idtipo=3"
        If mRs_Fua.RecordCount > 0 Then
           mRs_Fua.MoveFirst
           Do While Not mRs_Fua.EOF
              mrs_Diagnosticos.MoveFirst
              Do While Not mrs_Diagnosticos.EOF
                 If mrs_Diagnosticos.Fields!idDiagnostico = mRs_Fua!Item Then
                     mrs_Diagnosticos.Fields!FUA = mRs_Fua!idFuaCorrelativo
                     If Not IsNull(mRs_Fua!FuaCodigoPrestacion) Then
                        mrs_Diagnosticos.Fields!FuaCodigoPrestacion = mRs_Fua!FuaCodigoPrestacion
                     End If
                 End If
                 mrs_Diagnosticos.MoveNext
              Loop
              mRs_Fua.MoveNext
           Loop
           mrs_Diagnosticos.MoveFirst
        End If
    End If
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion_actual = lValue
End Property
Property Let Consultorio(lValue As String)
   ml_Consultorio = lValue
End Property
Property Let AScorrelativo(lValue As Long)
   GenerarRs
   ml_AScorrelativo = lValue
   If mrs_Diagnosticos.RecordCount > 0 Then
       mrs_Diagnosticos.MoveFirst
       Do While Not mrs_Diagnosticos.EOF
          mrs_Diagnosticos.Fields!Consultorio = ml_Consultorio
          mrs_Diagnosticos.Fields!idCuentaAtencion = ml_idCuentaAtencion_actual
          If IsNull(mrs_Diagnosticos.Fields!FUA) Then
             mrs_Diagnosticos.Fields!FUA = 1
          End If
          mrs_Diagnosticos.Update
          mrs_Diagnosticos.MoveNext
       Loop
   End If
   
   Set rsFUAequiv = mo_AdminAdmision.ServiciosAtenSimultaneaFUAequivXups(ml_ups)
   lnFUAequiv = rsFUAequiv.RecordCount
   Dim lnFor As Integer
   If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Header.Caption = "CPrestación"
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Hidden = False
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Width = 500
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Header.Appearance.ForeColor = vbWhite
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Header.Appearance.BackColor = vbRed
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Header.Appearance.Font.Bold = True
        If ml_AScorrelativo = 0 Then
           grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 1000
        End If
        grdDiagnosticos.Bands(0).Columns("fua").Header.Appearance.ForeColor = vbWhite
        grdDiagnosticos.Bands(0).Columns("fua").Header.Appearance.BackColor = vbRed
        grdDiagnosticos.Bands(0).Columns("fua").Header.Appearance.Font.Bold = True
        grdDiagnosticos.Bands(0).Columns("Fua").Width = 700
        On Error Resume Next
        With grdDiagnosticos.ValueLists.Add("FuaList").ValueListItems
            For lnFor = 1 To 20
                .Add lnFor, "N° " & Trim(Str(lnFor))
            Next
        End With
        grdDiagnosticos.Bands(0).Columns("Fua").ValueList = "FuaList"
        grdDiagnosticos.Bands(0).Columns("fua").Hidden = False
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Hidden = False
   Else
        grdDiagnosticos.Bands(0).Columns("fua").Hidden = True
        grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Hidden = True
        lnFUAequiv = 0
   End If
   
   If ml_AScorrelativo > 0 Then
        
        grdDiagnosticos.Bands(0).Columns("consultorio").Hidden = False
        grdDiagnosticos.Bands(0).Columns("consultorio").Header.Caption = "UPS"
        grdDiagnosticos.Bands(0).Columns("consultorio").Header.Appearance.Font.Bold = True
        '
        lblConsultorioActual.Visible = False
        cmbConsultorios.Visible = True
        Dim rsDiagnosticos As New Recordset
        Dim oConexion As New Connection
        Set rsConsultorios = mo_AdminAdmision.ServiciosAtenSimultaneaMovXcorrelativo(ml_AScorrelativo, True)
        mo_cmbConsultorios.BoundColumn = "IdCuentaAtencion"
        mo_cmbConsultorios.ListField = "Consultorio"
        Set mo_cmbConsultorios.RowSource = rsConsultorios
        mo_cmbConsultorios.BoundText = Trim(Str(ml_idCuentaAtencion_actual))
        'agregar Dx ya grabados de los Otros Consultorios
        If rsConsultorios.RecordCount > 0 Then
            oConexion.CommandTimeout = 300
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighentidades.CadenaConexion
            rsConsultorios.MoveFirst
            Do While Not rsConsultorios.EOF
                If ml_idCuentaAtencion_actual <> rsConsultorios!idCuentaAtencion Then
                    Set rsDiagnosticos = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsConsultorios!idAtencion, _
                                                                                                    ml_TipoDiagnostico, oConexion)
                    Do While Not rsDiagnosticos.EOF
                        With mrs_Diagnosticos
                            .AddNew
                            .Fields!idTipoDiagnostico = rsDiagnosticos!idTipoDiagnostico
                            .Fields!DescripcionTipoDx = rsDiagnosticos!DescripcionTipoDx
                            .Fields!idDiagnostico = rsDiagnosticos!idDiagnostico
                            .Fields!CodigoCIE2004 = rsDiagnosticos!CodigoCIE2004
                            .Fields!Descripcion = rsDiagnosticos!Descripcion
                            .Fields!labConfHIS = rsDiagnosticos!labConfHIS
                            .Fields!Consultorio = rsConsultorios!Consultorio
                            .Fields!idCuentaAtencion = rsConsultorios!idCuentaAtencion
                            .Fields!FUA = 1
                            .Fields!Grupo = IIf(IsNull(rsDiagnosticos!grupoHIS), 0, rsDiagnosticos!grupoHIS)
                            .Fields!SubGrupo = IIf(IsNull(rsDiagnosticos!subgrupoHIS), 0, rsDiagnosticos!subgrupoHIS)
                            .Fields!IdServicio = rsConsultorios!idCuentaAtencion
                            .Update
                        End With
                        rsDiagnosticos.MoveNext
                    Loop
                    If mrs_Diagnosticos.RecordCount > 0 Then
                       mrs_Diagnosticos.MoveFirst
                    End If
                End If
                rsConsultorios.MoveNext
            Loop
            oConexion.Close
        End If
        Set oConexion = Nothing
        Set rsDiagnosticos = Nothing
        
        grdDiagnosticos.Bands(0).Columns("consultorio").Hidden = False
        grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 5200
  Else
        lblConsultorioActual.Caption = ml_Consultorio
        lblConsultorioActual.Left = UserControl.cmbIdTipoDiagnostico.Left
        lblConsultorioActual.Top = cmbConsultorios.Top
        lblConsultorioActual.Width = cmbConsultorios.Width
        lblConsultorioActual.Visible = True
        cmbConsultorios.Visible = False
        
        grdDiagnosticos.Bands(0).Columns("consultorio").Hidden = True
        If wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
           grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 5200
        Else
           grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 8000
        End If
        
   End If
End Property
Property Let IdListBarItem(lValue As Long)
   ml_IdListBarItem = lValue
End Property
Property Let idAtencion(lValue As Long)
   ml_idAtencion = lValue
End Property
Property Get idAtencion() As Long
   idAtencion = ml_idAtencion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let TituloFrame(sValue As String)
   fraDiagnostico.Caption = sValue
End Property
Property Let TipoDiagnostico(lValue As sghTiposDiagnostico)
    ml_TipoDiagnostico = lValue
End Property
Property Get TipoDiagnostico() As sghTiposDiagnostico
   TipoDiagnostico = ml_TipoDiagnostico
End Property


Property Let SexoPaciente(lValue As Integer)
    ml_SexoPaciente = lValue
End Property
Property Let EdadPaciente(lValue As Long)
    ml_EdadPaciente = lValue
End Property
'mgaray201410c
Property Get rsDiagnosticos() As ADODB.Recordset
   Set rsDiagnosticos = mrs_Diagnosticos
End Property

Property Let BotonAgregarEnabled(bValue As Boolean)
    UserControl.btnAgregarDx.Enabled = False
End Property
Property Let BotonQuitarEnabled(bValue As Boolean)
    UserControl.btnQuitarDx.Enabled = False
End Property





Private Sub cmbConsultorios_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbConsultorios
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdTipoDiagnostico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoDiagnostico
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub cmbIdTipoDiagnostico_LostFocus()
   If cmbIdTipoDiagnostico.Text <> "" Then
        Dim lIdTipoDiagnostico As Long
        lIdTipoDiagnostico = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarIdPorCodigoYClasificacion(UCase(Split(cmbIdTipoDiagnostico.Text, " = ")(0)), ml_TipoDiagnostico)
        mo_cmbIdTipoDiagnostico.BoundText = lIdTipoDiagnostico
   End If
End Sub


Sub ChequeaDxVSpacienteEdadSexo()
         Dim lbContinuar As Boolean
         lbContinuar = True
         If UserControl.txtIdDiagnostico.Text <> "" And lbContinuar = True Then
             Dim oDODiagnostico As DODiagnostico
             If ml_IdListBarItem = sghOpcionGalenHos.sghRegistroAtencionCE Then
                Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(UserControl.txtIdDiagnostico.Text, False)
             Else
                Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorCodigoCIE2004(UserControl.txtIdDiagnostico.Text, True)
             End If
             If Not oDODiagnostico Is Nothing Then
                 UserControl.txtIdDiagnostico.Tag = oDODiagnostico.idDiagnostico
                 UserControl.lblDescripcionDx = oDODiagnostico.Descripcion
                 If oDODiagnostico.Restriccion Then
                     If oDODiagnostico.idTipoSexo <> 0 Then
                         If ml_SexoPaciente <> oDODiagnostico.idTipoSexo Then
                             MsgBox "El diagnóstico no corresponde al sexo del paciente", vbInformation, "Validación paciente"
                             UserControl.txtIdDiagnostico.Tag = ""
                             UserControl.lblDescripcionDx = ""
                             Exit Sub
                         End If
                         If ml_SexoPaciente = 1 And oDODiagnostico.Gestacion = True Then
                             MsgBox "El diagnóstico de gestación no corresponde al sexo del paciente ", vbInformation, "Validación paciente"
                             UserControl.txtIdDiagnostico.Tag = ""
                             UserControl.lblDescripcionDx = ""
                             Exit Sub
                         End If
                     End If
                     If ml_EdadPaciente < 3650 And oDODiagnostico.Gestacion = True Then
                         MsgBox "El diagnóstico de gestación no corresponde a la edad del paciente ", vbInformation, "Validación paciente"
                         UserControl.txtIdDiagnostico.Tag = ""
                         UserControl.lblDescripcionDx = ""
                         Exit Sub
                     End If
                     If (ml_EdadPaciente > oDODiagnostico.EdadMaxDias) Or (ml_EdadPaciente < oDODiagnostico.EdadMinDias) Then
                         MsgBox "El diagnóstico no corresponde a la edad del paciente (Edad mínima " & oDODiagnostico.EdadMinDias & " días - Edad máxima " & oDODiagnostico.EdadMaxDias & " días)", vbInformation, "Validación paciente"
                         UserControl.txtIdDiagnostico.Tag = ""
                         UserControl.lblDescripcionDx = ""
                         Exit Sub
                     End If
                 End If
             Else
                 UserControl.txtIdDiagnostico.Tag = ""
                 UserControl.lblDescripcionDx = ""
             End If
       End If
       mo_Formulario.MarcarComoVacio txtIdDiagnostico
End Sub

Public Function Inicializar()
    GenerarRecordsetTemporal
    mo_Formulario.HabilitarDeshabilitar UserControl.lblDescripcionDx, False
    Set mo_cmbIdTipoDiagnostico.MiComboBox = cmbIdTipoDiagnostico
    Set mo_cmbConsultorios.MiComboBox = cmbConsultorios
    
    On Error Resume Next
    txtIdDiagnostico.SetFocus
End Function

Public Sub FocusEnDx()
    On Error Resume Next
    txtIdDiagnostico.SetFocus
End Sub




Private Sub cmbLabHis_Click()
     Dim lnNuevo As Boolean
     lnNuevo = True
     If mrs_Labs.RecordCount > 0 Then
        mrs_Labs.MoveFirst
        mrs_Labs.Find "lab='" & Right(Trim(cmbLabHis.Text), 3) & "'"
        If Not mrs_Labs.EOF Then
           lnNuevo = False
        End If
     End If
     If lnNuevo = True Then
        mrs_Labs.AddNew
        mrs_Labs.Fields!lab = Right(Trim(cmbLabHis.Text), 3)
        mrs_Labs.Update
     End If
     grdLabHIS.Caption = ""
     Set UserControl.grdLabHIS.DataSource = mrs_Labs
     On Error Resume Next
     mrs_Labs.MoveFirst
End Sub









Private Sub cmbLabHis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbLabHis_Click
    End If
End Sub

Private Sub grdLabHIS_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdLabHIS.Bands(0).Columns("lab").Width = 500
End Sub

Private Sub UserControl_Resize()
    Dim lnPixelqAumentaron As Long
    UserControl.lblDescripcionDx.Width = UserControl.Width - 3070 - fraLab.Width
    fraLab.Left = lblDescripcionDx.Left + lblDescripcionDx.Width + 50
    fraDiagnostico.Width = UserControl.Width - 10
    fraDiagnostico.Height = UserControl.Height - 20
    UserControl.grdDiagnosticos.Width = fraDiagnostico.Width - 200 ' UserControl.Width - 20
    UserControl.grdDiagnosticos.Height = fraDiagnostico.Height - fraLab.Height - 300 'UserControl.Height - 1100
    
End Sub



Public Sub ConfigurarComboBoxes()
Dim sMensaje As String

        '
        mo_cmbIdTipoDiagnostico.BoundColumn = "IdSubclasificacionDx"
        mo_cmbIdTipoDiagnostico.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoDiagnostico.MiComboBox = cmbIdTipoDiagnostico
        Select Case ml_TipoDiagnostico
        Case sghAtencionConsultaExterna
            Set oRsTipoDx = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
            Set mo_cmbIdTipoDiagnostico.RowSource = oRsTipoDx ' mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxConsultaExterna
        Case sghHospitalizacionIngreso
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospIngreso
        Case sghHospitalizacionEgreso
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospEgreso
        Case sghHospitalizacionMortalidad
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospMortalidad
        Case sghHospitalizacionNacimiento
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospMuerteFetal
        Case sghHospitalizacionComplicaciones
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxHospComplicaciones
        Case sghInterconsultas
            Set mo_cmbIdTipoDiagnostico.RowSource = mo_AdminServiciosComunes.SubclasificacionDiagnosticosSeleccionarDxInterconsultas
        End Select
        
        Dim rsTipoDiagnostico As New Recordset
        Set rsTipoDiagnostico = mo_cmbIdTipoDiagnostico.RowSource
        Select Case rsTipoDiagnostico.RecordCount
        Case 0
            lblDiagnostico.Visible = False
            cmbIdTipoDiagnostico.Visible = False
        Case 1
        Case 2
        End Select
        
        sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError

End Sub
Private Sub btnBusquedaDiagnostico_Click()
    BusquedaDx ""
End Sub

Sub BusquedaDx(lcCodigoDx As String)
    Dim oBusqueda As New SIGHNegocios.BuscaDiagnosticos
    Dim oDODiagnostico As DODiagnostico
    If ml_IdListBarItem = sghOpcionGalenHos.sghRegistroAtencionCE Then
       oBusqueda.SoloMuestraDxGalenHos = False
    Else
       oBusqueda.SoloMuestraDxGalenHos = True
    End If
    oBusqueda.CodigoDx = lcCodigoDx
    oBusqueda.MostrarFormulario
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDODiagnostico = mo_AdminServiciosComunes.DiagnosticosSeleccionarPorId(oBusqueda.idRegistroSeleccionado)
        If Not oDODiagnostico Is Nothing Then
            UserControl.txtIdDiagnostico.Text = oDODiagnostico.CodigoCIE2004
            UserControl.txtIdDiagnostico.Tag = oDODiagnostico.idDiagnostico
            UserControl.lblDescripcionDx = oDODiagnostico.Descripcion
            On Error Resume Next
            cmbIdTipoDiagnostico.SetFocus
        Else
            UserControl.txtIdDiagnostico.Text = ""
            UserControl.txtIdDiagnostico.Tag = ""
            UserControl.lblDescripcionDx = ""
        End If
    Else
        UserControl.txtIdDiagnostico.Text = ""
        UserControl.txtIdDiagnostico.Tag = ""
        UserControl.lblDescripcionDx = ""
    End If
    Set oBusqueda = Nothing
End Sub


Private Sub grdDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdDiagnosticos.Bands(0).Columns("IdTipoDiagnostico").Hidden = True
    grdDiagnosticos.Bands(0).Columns("IdCuentaAtencion").Hidden = True
    grdDiagnosticos.Bands(0).Columns("idServicio").Hidden = True
    grdDiagnosticos.Bands(0).Columns("labConfHIS").Header.Caption = "Lab (HIS)"
    grdDiagnosticos.Bands(0).Columns("consultorio").Width = 1700
    grdDiagnosticos.Bands(0).Columns("consultorio").Header.Appearance.ForeColor = vbWhite
    grdDiagnosticos.Bands(0).Columns("consultorio").Header.Appearance.BackColor = vbRed
    grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Width = 1700
    grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Header.Appearance.ForeColor = vbWhite
    grdDiagnosticos.Bands(0).Columns("FuaCodigoPrestacion").Header.Appearance.BackColor = vbRed

    grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Header.Caption = "Tipo diagnóstico"
    grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Width = 800
    grdDiagnosticos.Bands(0).Columns("DescripcionTipoDx").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
    
    grdDiagnosticos.Bands(0).Columns("IdDiagnostico").Hidden = True
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Header.Caption = "CIE"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Width = 1000
    grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
    
    grdDiagnosticos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdDiagnosticos.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
    
    grdDiagnosticos.Bands(0).Columns("labConfHIS").Activation = ssActivationActivateNoEdit 'Actualizado 25092014
    grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 5200
    grdDiagnosticos.Bands(0).Columns("grupo").Width = 300
    grdDiagnosticos.Bands(0).Columns("subgrupo").Width = 300
   '
    
    mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, sighentidades.GrillaConFilasBicolor
End Sub
Private Sub txtIdDiagnostico_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdDiagnostico
    If KeyCode = vbKeyF1 Then
        btnBusquedaDiagnostico_Click
    End If
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtIdDiagnostico_LostFocus()
   If Len(txtIdDiagnostico.Text) > 0 And lblDescripcionDx.Text = "" Then
      BusquedaDx txtIdDiagnostico.Text
   End If
   UserControl.txtIdDiagnostico.Text = UCase(UserControl.txtIdDiagnostico.Text)
End Sub

Private Sub txtIdDiagnostico_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
        If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub CargarDatosDeDiagnosticos(oConexion As Connection)
Dim rsDiagnosticos As New Recordset

    Set rsDiagnosticos = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(ml_idAtencion, ml_TipoDiagnostico, oConexion)
    Do While Not rsDiagnosticos.EOF
        With mrs_Diagnosticos
            .AddNew
            .Fields!idTipoDiagnostico = rsDiagnosticos!idTipoDiagnostico
            .Fields!DescripcionTipoDx = rsDiagnosticos!DescripcionTipoDx
            .Fields!idDiagnostico = rsDiagnosticos!idDiagnostico
            .Fields!CodigoCIE2004 = rsDiagnosticos!CodigoCIE2004
            .Fields!Descripcion = rsDiagnosticos!Descripcion
            .Fields!labConfHIS = rsDiagnosticos!labConfHIS
            .Fields!Grupo = IIf(IsNull(rsDiagnosticos!grupoHIS), 0, rsDiagnosticos!grupoHIS)
            .Fields!SubGrupo = IIf(IsNull(rsDiagnosticos!subgrupoHIS), 0, rsDiagnosticos!subgrupoHIS)
            .Fields!IdServicio = ml_idCuentaAtencion_actual
            'If ml_AScorrelativo > 0 Then
               .Fields!FUA = 1
            'End If
        End With
        rsDiagnosticos.MoveNext
    Loop
'    On Error Resume Next
    If mrs_Diagnosticos.RecordCount > 0 Then
       mrs_Diagnosticos.MoveFirst
    End If
    Set grdDiagnosticos.DataSource = mrs_Diagnosticos
    
    Set rsDiagnosticos = Nothing
End Sub


Sub CargarDiagnosticosAlObjetoDatos(oDiagnosticos As Collection)
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS DIAGNOSTICOS
    '---------------------------------------------------------------------------------
    Dim oRow As SSRow
    Dim oDiagnostico As DOAtencionDiagnostico
    Set oRow = UserControl.grdDiagnosticos.GetRow(ssChildRowFirst)
    
    If Not oRow Is Nothing Then
        'Para el primero
        If ml_idCuentaAtencion_actual = oRow.Cells("idCuentaAtencion").Value Then
            Set oDiagnostico = New DOAtencionDiagnostico
            oDiagnostico.IdAtencionDiagnostico = 0
            oDiagnostico.idAtencion = ml_idAtencion
            oDiagnostico.idDiagnostico = oRow.Cells("IdDiagnostico").Value
            oDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
            oDiagnostico.IdSubclasificacionDx = IIf(IsNull(oRow.Cells("IdTipoDiagnostico").Value), 0, oRow.Cells("IdTipoDiagnostico").Value)
            oDiagnostico.IdUsuarioAuditoria = ml_idUsuario
            oDiagnostico.labConfHIS = IIf(IsNull(oRow.Cells("labConfHIS").Value), "", oRow.Cells("labConfHIS").Value)
            oDiagnostico.grupoHIS = oRow.Cells("grupo").Value
            oDiagnostico.subgrupoHIS = oRow.Cells("subGrupo").Value
            oDiagnosticos.Add oDiagnostico
        End If
        'Para los siguientes
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If ml_idCuentaAtencion_actual = oRow.Cells("idCuentaAtencion").Value Then
                Set oDiagnostico = New DOAtencionDiagnostico
                oDiagnostico.IdAtencionDiagnostico = 0
                oDiagnostico.idAtencion = ml_idAtencion
                oDiagnostico.idDiagnostico = oRow.Cells("IdDiagnostico").Value
                oDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
                oDiagnostico.IdSubclasificacionDx = IIf(IsNull(oRow.Cells("IdTipoDiagnostico").Value), 0, oRow.Cells("IdTipoDiagnostico").Value)
                oDiagnostico.IdUsuarioAuditoria = ml_idUsuario
                oDiagnostico.labConfHIS = IIf(IsNull(oRow.Cells("labConfHIS").Value), "", oRow.Cells("labConfHIS").Value)
                oDiagnostico.grupoHIS = oRow.Cells("grupo").Value
                oDiagnostico.subgrupoHIS = oRow.Cells("subGrupo").Value
                oDiagnosticos.Add oDiagnostico
            End If
        Loop
    End If
End Sub

Sub GenerarRs()
    grdLabHIS.Caption = ""
    If mrs_Labs.State = 1 Then Set mrs_Labs = Nothing
    With mrs_Labs
          .Fields.Append "lab", adVarChar, 3, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set UserControl.grdLabHIS.DataSource = mrs_Labs
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdLabHIS, sighentidades.GrillaConFilasBicolor
End Sub

Public Sub GenerarRecordsetTemporal()
    GenerarRs
    '
    If mrs_Diagnosticos.State = 1 Then Set mrs_Diagnosticos = Nothing
    With mrs_Diagnosticos
          .Fields.Append "IdCuentaAtencion", adInteger
          .Fields.Append "IdTipoDiagnostico", adInteger, 4, adFldIsNullable
          .Fields.Append "DescripcionTipoDx", adVarChar, 100, adFldIsNullable
          .Fields.Append "IdDiagnostico", adInteger
          .Fields.Append "CodigoCIE2004", adVarChar, 10
          .Fields.Append "Descripcion", adVarChar, 255
          .Fields.Append "labConfHIS", adVarChar, 3, adFldIsNullable
          .Fields.Append "Fua", adInteger
          .Fields.Append "FuaCodigoPrestacion", adVarChar, 3, adFldIsNullable
          .Fields.Append "Consultorio", adVarChar, 100, adFldIsNullable
          .Fields.Append "idServicio", adInteger
          .Fields.Append "grupo", adInteger
          .Fields.Append "subgrupo", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set UserControl.grdDiagnosticos.DataSource = mrs_Diagnosticos
    mo_Apariencia.ConfigurarFilasBiColores UserControl.grdDiagnosticos, sighentidades.GrillaConFilasBicolor
End Sub

Sub LimpiarDatos()
    On Error GoTo errLimp
    With mrs_Diagnosticos
       If .RecordCount > 0 Then
          .MoveFirst
          Do While Not .EOF
             .Delete
             .Update
             .MoveNext
          Loop
       End If
    End With
errLimp:
End Sub

Private Sub btnAgregarDx_Click()
    Dim lbContinuarLab As Boolean, lcLab As String, lcFuaCodigoPrestacion As String, lcTipoDx As String
    Dim lcSql As String, lnRegDx As Long, lnFua As Long, lnEdad As Integer
    
    ChequeaDxVSpacienteEdadSexo
    
    If UserControl.txtIdDiagnostico.Text = "" Then
        MsgBox "Por favor ingrese el diagnóstico", vbInformation, "Diagnósticos"
        Exit Sub
    End If
    
    If UserControl.txtIdDiagnostico.Tag = "" Then
        MsgBox "Por favor ingrese un diagnóstico válido", vbInformation, "Diagnósticos"
        Exit Sub
    End If
    
    If cmbIdTipoDiagnostico.Visible = True Then
        If UserControl.cmbIdTipoDiagnostico.Text = "" Then
            MsgBox "Por favor ingrese el tipo de diagnóstico", vbInformation, "Diagnósticos"
            Exit Sub
        End If
    End If
    
    '***************daniel barrantes**************
    '***************Valida Diagnosticos REPETIDOS
    '***************
    lnFua = 1
    lnRegDx = mrs_Diagnosticos.RecordCount
    If lnRegDx > 0 Then
        mrs_Diagnosticos.MoveFirst
        Do While Not mrs_Diagnosticos.EOF
        
            'Yamill Palomino
            If ml_IdListBarItem = sghOpcionGalenHos.sghAdmisionEmergencia Or ml_IdListBarItem = sghOpcionGalenHos.sghAdmisionHospitalizacion Then
                If txtIdDiagnostico.Tag = mrs_Diagnosticos!idDiagnostico Then
                    MsgBox "El diagnóstico ya fue agregado al listado", vbInformation, "Admision"
                    Exit Sub
                End If
            Else
            'Actualizado 29092014
    '            If txtIdDiagnostico.Tag = mrs_Diagnosticos!IdDiagnostico And Trim(cmbLabHis.Text) = mrs_Diagnosticos!labConfHIS And mrs_Diagnosticos!idTipoDiagnostico = Val(mo_cmbIdTipoDiagnostico.BoundText) Then
                If txtIdDiagnostico.Tag = mrs_Diagnosticos!idDiagnostico And Trim(cmbLabHis.Text) = IIf(IsNull(mrs_Diagnosticos!labConfHIS) = True, "", Trim(mrs_Diagnosticos!labConfHIS)) Then
                    If Trim(cmbLabHis.Text) = "" Then
                        MsgBox "El diagnóstico ya fué agregado al listado", vbInformation, "Admisión"
                    Else
                        MsgBox "El diagnóstico con el mismo codigo lab ya fué registrado", vbInformation, "Admisión"
                    End If
                    'Exit Sub
                End If
            End If
            If mrs_Diagnosticos!FUA > lnFua Then
               lnFua = mrs_Diagnosticos!FUA
            End If
            mrs_Diagnosticos.MoveNext
        Loop
    End If
    '
    If mrs_Labs.RecordCount > 0 Then
       mrs_Labs.MoveFirst
    End If
    lbContinuarLab = True
    Do While lbContinuarLab
        lcLab = ""
        If mrs_Labs.RecordCount = 0 Then
           lbContinuarLab = False
        Else
           lcLab = mrs_Labs!lab
           mrs_Labs.MoveNext
           If mrs_Labs.EOF Then
              lbContinuarLab = False
           End If
        End If
        'chequea si tiene CODIGO PRESTACION automatico
        lnRegDx = mrs_Diagnosticos.RecordCount
        lcFuaCodigoPrestacion = ""
        If lnFUAequiv > 0 And wxParametro302 = "S" And ml_IdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
           lcTipoDx = IIf(Val(mo_cmbIdTipoDiagnostico.BoundText) = 101, "P", IIf(Val(mo_cmbIdTipoDiagnostico.BoundText) = 102, "D", "R"))
           lcSql = "DxCodigo='" & Trim(UserControl.txtIdDiagnostico.Text) & "' and DxTipo='" & lcTipoDx & "'" & _
                               " and (PesoKgMenor<=" & ml_PesoKg & " and PesoKgMayor>=" & ml_PesoKg & ") "
           rsFUAequiv.Filter = lcSql
           If rsFUAequiv.RecordCount > 0 Then
                rsFUAequiv.MoveFirst
                Do While Not rsFUAequiv.EOF
                   Select Case rsFUAequiv!IdTipoEdad
                   Case 1
                       lnEdad = DateDiff("yyyy", ml_FechaNacimiento, ml_FechaAtencion)
                   Case 2
                       lnEdad = DateDiff("m", ml_FechaNacimiento, ml_FechaAtencion)
                   Case 3
                       lnEdad = DateDiff("d", ml_FechaNacimiento, ml_FechaAtencion)
                   End Select
                   If lnEdad >= rsFUAequiv!EdadInicio And lnEdad <= rsFUAequiv!EdadFinal Then
                      If lnRegDx = 0 Then
                         lcFuaCodigoPrestacion = rsFUAequiv!FuaCodigoPrestacion
                      Else
                         mrs_Diagnosticos.Find "fuaCodigoPrestacion='" & rsFUAequiv!FuaCodigoPrestacion & "'"
                         If mrs_Diagnosticos.EOF Then
                            lcFuaCodigoPrestacion = rsFUAequiv!FuaCodigoPrestacion
                            lnFua = lnFua + 1
                            Exit Do
                         End If
                      End If
                   End If
                   rsFUAequiv.MoveNext
                Loop
           End If
        End If
        '
        With mrs_Diagnosticos
            .AddNew
            .Fields!idDiagnostico = Val(UserControl.txtIdDiagnostico.Tag)
            .Fields!CodigoCIE2004 = UserControl.txtIdDiagnostico.Text
            .Fields!Descripcion = UserControl.lblDescripcionDx
            .Fields!idTipoDiagnostico = Val(mo_cmbIdTipoDiagnostico.BoundText)
            .Fields!DescripcionTipoDx = UserControl.cmbIdTipoDiagnostico.Text
            .Fields!labConfHIS = lcLab
            If ml_AScorrelativo = 0 Then
                .Fields!Consultorio = ml_Consultorio
                .Fields!idCuentaAtencion = ml_idCuentaAtencion_actual
                .Fields!IdServicio = ml_IdServicio
            Else
                .Fields!Consultorio = cmbConsultorios.Text
                .Fields!idCuentaAtencion = Val(mo_cmbConsultorios.BoundText)
                .Fields!IdServicio = Val(mo_cmbConsultorios.BoundText)
            End If
            .Fields!FUA = lnFua
            .Fields!FuaCodigoPrestacion = lcFuaCodigoPrestacion
            .Update
        End With
        sighentidades.ParaAuditoriaPorCadaDato sghAudGrabaRegEdit, "+Dx: " & Trim(UserControl.txtIdDiagnostico.Text)
    Loop
    '
    GenerarRs
    '
    txtIdDiagnostico.Tag = ""
    txtIdDiagnostico.Text = ""
    lblDescripcionDx = ""
    mo_cmbIdTipoDiagnostico.BoundText = ""
'    cmbIdTipoDiagnostico.Text = ""
    cmbIdTipoDiagnostico.ListIndex = -1
    cmbLabHis.Text = ""
    
    On Error Resume Next
    mrs_Diagnosticos.MoveFirst
    'txtIdDiagnostico.SetFocus
    RaiseEvent SePresionoTeclaEspecial(vbKeyTab)
    If ml_IdListBarItem = sghOpcionGalenHos.sghRegistroAtencionCE Then  'CE-Reg.Dx
       FocusEnDx
    End If

End Sub

'Actualizado 15102014
Private Sub btnQuitarDx_Click()
'    On Error Resume Next
'    With mrs_Diagnosticos
'        If Not .EOF And Not .BOF Then
'           .Delete
'           .Update
'        End If
'    End With
'     Set UserControl.grdDiagnosticos.DataSource = mrs_Diagnosticos

    EliminarDiagnosticoSeleccionado
End Sub

Public Sub EditaLabConfHIS()
    'cmbLabHis.Visible = True
    'lblLabConfHIS.Visible = True
    grdDiagnosticos.Bands(0).Columns("labConfHIS").Width = 1000
    grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 1000
    'debb-9-2-211
    Set cmbLabHis.ListSource = mo_AdminServiciosComunes.DevuelveHIS_SITUACIOporDescripcion()
    'debb-9-2-211
End Sub

Public Sub TipoDxDefault(lcValorDefault As String)            'debb-06-03-2012
    If cmbIdTipoDiagnostico.Text = "" Then
       mo_cmbIdTipoDiagnostico.BoundText = lcValorDefault
    End If
End Sub

'Actualizado 25092014
Private Sub grdDiagnosticos_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
    EliminarDiagnosticoSeleccionado
End Sub

Public Sub EliminarDiagnosticoSeleccionado()
    If MsgBox("¿Desea eliminar el diagnóstico seleccionado?", vbYesNo, "Eliminar diagnósticos") = vbYes Then
        On Error Resume Next
        If mrs_Diagnosticos.RecordCount > 0 Then
            With mrs_Diagnosticos
                If Not .EOF And Not .BOF Then
                   .Delete
                   .Update
                End If
            End With
        End If
        Set grdDiagnosticos.DataSource = mrs_Diagnosticos
        If mrs_Diagnosticos.RecordCount > 0 Then
           mrs_Diagnosticos.MoveFirst
        End If
        
    End If
End Sub

'mgaray20141008
Public Function DeshabilitarEdicionDatos() As Boolean
    fraDiagnostico.Enabled = False
    grdDiagnosticos.Enabled = False
    btnBusquedaDiagnostico.Enabled = False
End Function

Public Function HabilitarEdicionDatos() As Boolean
    fraDiagnostico.Enabled = True
    grdDiagnosticos.Enabled = True
    btnBusquedaDiagnostico.Enabled = True
End Function

'A.Yañez*************************************
Public Function limpiacampos() As Boolean
     txtIdDiagnostico.Text = ""
     lblDescripcionDx.Text = ""
End Function
'*********************************************


'debb-23/02/2015
Sub CargarDatosDeDiagnosticosEmergCE(oConexion As Connection, lnTipoDiagnostico As Long)
Dim rsDiagnosticos As New Recordset

    Set rsDiagnosticos = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(ml_idAtencion, lnTipoDiagnostico, oConexion)
    Do While Not rsDiagnosticos.EOF
        With mrs_Diagnosticos
            .AddNew
            .Fields!idTipoDiagnostico = rsDiagnosticos!idTipoDiagnostico
            .Fields!DescripcionTipoDx = rsDiagnosticos!DescripcionTipoDx
            .Fields!idDiagnostico = rsDiagnosticos!idDiagnostico
            .Fields!CodigoCIE2004 = rsDiagnosticos!CodigoCIE2004
            .Fields!Descripcion = rsDiagnosticos!Descripcion
            .Fields!labConfHIS = rsDiagnosticos!labConfHIS
        End With
        rsDiagnosticos.MoveNext
    Loop
    On Error Resume Next
    rsDiagnosticos.MoveFirst
    
    Set rsDiagnosticos = Nothing
End Sub

Sub CargarDiagnosticosAlObjetoDatosMenosCtaActual(oDiagnosticos As Collection)
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DE LOS DIAGNOSTICOS
    '---------------------------------------------------------------------------------
    Dim oRow As SSRow
    Dim oDiagnostico As DOAtencionDiagnostico
    Set oRow = UserControl.grdDiagnosticos.GetRow(ssChildRowFirst)
    
    If Not oRow Is Nothing Then
        'Para el primero
        If ml_idCuentaAtencion_actual <> oRow.Cells("idCuentaAtencion").Value Then
            rsConsultorios.MoveFirst
            rsConsultorios.Find "idCuentaAtencion=" & oRow.Cells("idCuentaAtencion").Value
            Set oDiagnostico = New DOAtencionDiagnostico
            oDiagnostico.IdAtencionDiagnostico = 0
            oDiagnostico.idAtencion = rsConsultorios!idAtencion
            oDiagnostico.idDiagnostico = oRow.Cells("IdDiagnostico").Value
            oDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
            oDiagnostico.IdSubclasificacionDx = IIf(IsNull(oRow.Cells("IdTipoDiagnostico").Value), 0, oRow.Cells("IdTipoDiagnostico").Value)
            oDiagnostico.IdUsuarioAuditoria = ml_idUsuario
            oDiagnostico.labConfHIS = IIf(IsNull(oRow.Cells("labConfHIS").Value), "", oRow.Cells("labConfHIS").Value)
            oDiagnostico.grupoHIS = oRow.Cells("grupo").Value
            oDiagnostico.subgrupoHIS = oRow.Cells("subGrupo").Value
            oDiagnosticos.Add oDiagnostico
        End If
        'Para los siguientes
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If ml_idCuentaAtencion_actual <> oRow.Cells("idCuentaAtencion").Value Then
                rsConsultorios.MoveFirst
                rsConsultorios.Find "idCuentaAtencion=" & oRow.Cells("idCuentaAtencion").Value
                Set oDiagnostico = New DOAtencionDiagnostico
                oDiagnostico.IdAtencionDiagnostico = 0
                oDiagnostico.idAtencion = rsConsultorios!idAtencion
                oDiagnostico.idDiagnostico = oRow.Cells("IdDiagnostico").Value
                oDiagnostico.IdClasificacionDx = ml_TipoDiagnostico
                oDiagnostico.IdSubclasificacionDx = IIf(IsNull(oRow.Cells("IdTipoDiagnostico").Value), 0, oRow.Cells("IdTipoDiagnostico").Value)
                oDiagnostico.IdUsuarioAuditoria = ml_idUsuario
                oDiagnostico.labConfHIS = IIf(IsNull(oRow.Cells("labConfHIS").Value), "", oRow.Cells("labConfHIS").Value)
                oDiagnostico.grupoHIS = oRow.Cells("grupo").Value
                oDiagnostico.subgrupoHIS = oRow.Cells("subGrupo").Value
                oDiagnosticos.Add oDiagnostico
            End If
        Loop
    End If
End Sub


Public Function DevuelveDx() As Recordset
    mrs_Diagnosticos.Filter = ""
    Set DevuelveDx = mrs_Diagnosticos
End Function

Public Sub EliminaLosQueTienenGrupo()
        If mrs_Diagnosticos.RecordCount > 0 Then
           mrs_Diagnosticos.MoveFirst
           Do While Not mrs_Diagnosticos.EOF
              If mrs_Diagnosticos!Grupo > 0 Then
                 mrs_Diagnosticos.Delete
                 mrs_Diagnosticos.Update
              End If
              mrs_Diagnosticos.MoveNext
           Loop
        End If
End Sub
