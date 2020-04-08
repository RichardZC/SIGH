VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl ucDiagnosticosLista 
   ClientHeight    =   5592
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10116
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5592
   ScaleWidth      =   10116
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   75
      TabIndex        =   5
      Top             =   540
      Width           =   10005
      Begin VB.CheckBox chkFiltroIzq 
         Caption         =   "Filtro desde la IZQUIERDA"
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
         Left            =   6675
         TabIndex        =   10
         Top             =   540
         Width           =   2880
      End
      Begin VB.CheckBox chkAccess 
         Caption         =   "ACCESS"
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
         Left            =   8760
         TabIndex        =   9
         Top             =   195
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CheckBox chkPorCadaLetra 
         Caption         =   "Filtro x letra"
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
         Left            =   6675
         TabIndex        =   8
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5310
         Picture         =   "ucDiagnosticosLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   5340
         Picture         =   "ucDiagnosticosLista.ctx":2C49
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   510
         Width           =   1275
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   0
         Top             =   480
         Width           =   1065
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   1
         Top             =   480
         Width           =   3915
      End
      Begin VB.Label Label2 
         Caption         =   "Código           Descripción   "
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
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   5055
      End
   End
   Begin UltraGrid.SSUltraGrid grdDiagnosticos 
      Height          =   4050
      Left            =   75
      TabIndex        =   4
      Top             =   1515
      Width           =   10005
      _ExtentX        =   17653
      _ExtentY        =   7154
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
      Caption         =   "Lista de diagnósticos"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Diagnósticos (CIE-10)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15
      TabIndex        =   7
      Top             =   15
      Width           =   10080
   End
End
Attribute VB_Name = "ucDiagnosticosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para listar diagnósticos
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Public Event SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim ml_IdRegistroSeleccionado As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lbSoloMuestraDxGalenHos As Boolean
Dim lbUSAcodigoCIEsinPto As Boolean
Dim lbPorCadaLetra As Boolean
Dim oConexionMDB As New Connection
'mgaray09
Dim mb_mostrarSoloActivos As Boolean

Property Let USAcodigoCIEsinPto(lValue As Boolean)
    lbUSAcodigoCIEsinPto = lValue
End Property

Property Let SoloMuestraDxGalenHos(lValue As Boolean)
    lbSoloMuestraDxGalenHos = lValue
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdDiagnosticos.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdDiagnosticos.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

'mgaray09
Property Let MostrarSoloActivos(bValue As Boolean)
    mb_mostrarSoloActivos = bValue
End Property

Property Get MostrarSoloActivos() As Boolean
    MostrarSoloActivos = mb_mostrarSoloActivos
End Property

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda True
    Screen.MousePointer = vbDefault
End Sub

Function DiagnosticosFiltrarMDB(oDODiagnostico As doDiagnostico, lbSoloMuestraDxGalenHos As Boolean, _
                                lbUSAcodigoCIEsinPto As Boolean, Optional MostrarSoloActivos As Boolean = True) As Recordset
       Dim oRsDxMDB As New Recordset
       Dim sWhere As String, lcSql As String
       If lbUSAcodigoCIEsinPto = True Then
            'mgaray09
           lcSql = "select Iddiagnostico, codigoCIEsinPto, Descripcion,CodigoCIE10, CodigoCIE9, EsActivo, FechaInicioVigencia from Diagnosticos "
           If oDODiagnostico.codigoCIEsinPto <> "" Then
               sWhere = sWhere + " codigoCIEsinPto like '" + oDODiagnostico.codigoCIEsinPto + "*' and "
           End If
       Else
           'mgaray09
           lcSql = "select Iddiagnostico, CodigoCIE2004, Descripcion,CodigoCIE10, CodigoCIE9, EsActivo, FechaInicioVigencia from Diagnosticos "
           If oDODiagnostico.CodigoCIE2004 <> "" Then
              sWhere = sWhere + " CodigoCIE2004 like '" + oDODiagnostico.CodigoCIE2004 + "*' and "
           End If
       End If
       If oDODiagnostico.Descripcion <> "" And oDODiagnostico.Descripcion <> "%%" Then
            
            If Left(oDODiagnostico.Descripcion, 1) <> "%" Then
               oDODiagnostico.Descripcion = Trim(oDODiagnostico.Descripcion) & "*"
            End If
            sWhere = sWhere + " Descripcion like '" + oDODiagnostico.Descripcion + "' and "
       End If
       If lbSoloMuestraDxGalenHos = True Then
           sWhere = sWhere + " not (descripcionMINSA is null) and "
       End If
       'mgaray09
       If mb_mostrarSoloActivos = True Then
           sWhere = sWhere + " EsActivo = 1 and "
       End If
       If sWhere <> "" Then
            sWhere = " Where " & Left(sWhere, Len(sWhere) - 4)
       End If
       sWhere = sWhere + " order by  Descripcion,CodigoCIE2004 "
       lcSql = lcSql & sWhere
       oRsDxMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
       Set DiagnosticosFiltrarMDB = oRsDxMDB
End Function

Public Sub RealizarBusqueda(lbPasaAlGrid As Boolean)
Dim oDODiagnostico As New doDiagnostico
        If lbUSAcodigoCIEsinPto = True Then
           oDODiagnostico.codigoCIEsinPto = UserControl.txtCodigo
        Else
           oDODiagnostico.CodigoCIE2004 = UserControl.txtCodigo
        End If
        
        If chkAccess.value = 1 Then
           oDODiagnostico.Descripcion = "%" & Trim(UserControl.txtDescripcion) & "%"
           Set grdDiagnosticos.DataSource = DiagnosticosFiltrarMDB(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
        Else
            'Actualizado 2209
            If UserControl.txtDescripcion <> "" Then
                oDODiagnostico.Descripcion = "%" & Trim(UserControl.txtDescripcion) & "%"
            End If
'           'mgaray09
'           If mb_mostrarSoloActivos = False Then
'                Set grdDiagnosticos.DataSource = mo_AdminServiciosComunes.DiagnosticosFiltrar(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
'           Else
'                Set grdDiagnosticos.DataSource = mo_AdminServiciosComunes.DiagnosticosFiltrarSoloActivos(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
'           End If
           Dim oRsTmp As New Recordset
           Dim lcSql As String
           If mb_mostrarSoloActivos = False Then
                Set oRsTmp = mo_AdminServiciosComunes.DiagnosticosFiltrar(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
           Else
                Set oRsTmp = mo_AdminServiciosComunes.DiagnosticosFiltrarSoloActivos(oDODiagnostico, lbSoloMuestraDxGalenHos, lbUSAcodigoCIEsinPto)
           End If
           If chkFiltroIzq.value = 1 And UserControl.txtDescripcion <> "" Then
               lcSql = "Descripcion like '" & Trim(UserControl.txtDescripcion) & "%'"
               oRsTmp.Filter = lcSql
           End If
           Set grdDiagnosticos.DataSource = oRsTmp

        End If
        mo_Apariencia.ConfigurarFilasBiColores grdDiagnosticos, sighentidades.GrillaConFilasBicolor
        If Len(Trim(UserControl.txtDescripcion)) > 0 Then
           If lbPasaAlGrid = True Then
                On Error Resume Next
                grdDiagnosticos.SetFocus
           End If
        End If
End Sub




Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        UserControl.txtCodigo = ""
        UserControl.txtDescripcion = ""
End Sub

Private Sub chkAccess_Click()
   On Error Resume Next
   If chkAccess.value = 1 Then
      oConexionMDB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\archivos\tablas nuevas galenhos.mdb;Password=debb"
   Else
      oConexionMDB.Close
   End If
End Sub

Private Sub chkFiltroIzq_Click()
    If chkFiltroIzq.value = 1 Then
       sighentidades.BuscarSoloIzquierda = "1"
    Else
       sighentidades.BuscarSoloIzquierda = "0"
    End If
End Sub

Private Sub chkPorCadaLetra_Click()
    If chkPorCadaLetra.value = 1 Then
       lbPorCadaLetra = True
       
    Else
       lbPorCadaLetra = False
       
    End If
End Sub

Private Sub grdDiagnosticos_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdDiagnosticos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdDiagnostico")
 
End Sub

Private Sub grdDiagnosticos_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdDiagnosticos.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdDiagnostico")
    
End Sub


Private Sub grdDiagnosticos_DblClick()
    grdDiagnosticos_Click
    RaiseEvent SeleccionaRegistro(ml_IdRegistroSeleccionado)
End Sub

Private Sub grdDiagnosticos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdDiagnosticos.Bands(0).Columns("IdDiagnostico").Hidden = True
    
    If lbUSAcodigoCIEsinPto = True Then
        grdDiagnosticos.Bands(0).Columns("codigoCIEsinPto").Header.Caption = "CIE-10"
        grdDiagnosticos.Bands(0).Columns("codigoCIEsinPto").Width = 1000
    Else
        grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Header.Caption = "CIE-10"
        grdDiagnosticos.Bands(0).Columns("CodigoCIE2004").Width = 1000
    End If
    
    grdDiagnosticos.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdDiagnosticos.Bands(0).Columns("Descripcion").Width = 8400
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE10").Hidden = True
    grdDiagnosticos.Bands(0).Columns("CodigoCIE10").Header.Caption = "CIE10"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE10").Width = 1000
    
    grdDiagnosticos.Bands(0).Columns("CodigoCIE9").Header.Caption = "CIE-9"
    grdDiagnosticos.Bands(0).Columns("CodigoCIE9").Width = 1000
    
    grdDiagnosticos.Bands(0).Columns("EsActivo").Header.Caption = "Habilitado"
    grdDiagnosticos.Bands(0).Columns("EsActivo").Width = 500
    
    grdDiagnosticos.Bands(0).Columns("FechaInicioVigencia").Header.Caption = "F. Vigencia"
    grdDiagnosticos.Bands(0).Columns("FechaInicioVigencia").Width = 1100
    grdDiagnosticos.Bands(0).Columns("FechaInicioVigencia").CellAppearance.TextAlign = ssAlignCenter
    
    chkFiltroIzq.value = Val(sighentidades.BuscarSoloIzquierda)
End Sub

Private Sub grdDiagnosticos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdDiagnosticos_DblClick
    End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsCIE10(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(txtCodigo)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub



Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If lbPorCadaLetra = True Then
          RealizarBusqueda False
    End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   
   grdDiagnosticos.Width = fraBusqueda.Width
   grdDiagnosticos.Height = UserControl.Height - (lblNombre.Height + fraBusqueda.Height + 150)
   
End Sub

Property Let CodigoDx(lValue As String)
    txtCodigo.Text = lValue
    btnBuscar_Click
End Property

Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub


'debb2014b
Public Sub FocusEnDescripcion()
    On Error Resume Next
    If txtCodigo.Text = "" And txtDescripcion.Text = "" Then
       txtDescripcion.SetFocus
    Else
       grdDiagnosticos.SetFocus
    End If
End Sub
