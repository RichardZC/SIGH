VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form HerrActualizacionParametros 
   Caption         =   "Actualización de datos de tabla parametros"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmActualizacionParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6585
      Left            =   15
      TabIndex        =   3
      Top             =   45
      Width           =   11970
      Begin VB.CommandButton Ayuda 
         Caption         =   "..."
         Height          =   210
         Left            =   11730
         TabIndex        =   7
         Top             =   195
         Width           =   165
      End
      Begin VB.ComboBox cmbGrupo 
         Height          =   330
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   4440
      End
      Begin UltraGrid.SSUltraGrid grdParametros 
         Height          =   5745
         Left            =   0
         TabIndex        =   6
         Top             =   600
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   10134
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
         Caption         =   "Parametros"
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Grupos"
         Height          =   345
         Left            =   1080
         TabIndex        =   4
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   1
      Top             =   6600
      Width           =   11970
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FrmActualizacionParametros.frx":0CCA
         DownPicture     =   "FrmActualizacionParametros.frx":118E
         Height          =   700
         Left            =   6158
         Picture         =   "FrmActualizacionParametros.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FrmActualizacionParametros.frx":1B66
         DownPicture     =   "FrmActualizacionParametros.frx":1FC6
         Height          =   700
         Left            =   4600
         Picture         =   "FrmActualizacionParametros.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   135
         Width           =   1365
      End
   End
End
Attribute VB_Name = "HerrActualizacionParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Actualiza tabla parametros
'        Programado por: Benavides M
'        Fecha: Abril 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_cmbGrupo As New sighEntidades.ListaDespleglable
Dim mo_ReglasComunes As New ReglasComunes
Dim mo_Param As New DOPArametro
Dim sMensaje As String
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Teclado As New sighEntidades.Teclado
Dim ml_Grupo As String
Dim ml_TextoDelFiltro As String
Dim mrs_Parametros As New ADODB.Recordset
Dim oParametros As New Parametros


Private Sub Ayuda_Click()
        Dim oCambClave As New LoginActualizaClave
        oCambClave.idUsuario = sighEntidades.Usuario
        oCambClave.Show 1
        Set oCambClave = Nothing

End Sub

Private Sub cmbGrupo_Click()
    Dim orstemp1 As New Recordset
    GenerarRecordsetTemporal
    Set orstemp1 = mo_ReglasComunes.SeleccionarGrupoParametros(cmbGrupo.Text)
    If orstemp1.RecordCount > 0 Then
        orstemp1.MoveFirst
        Do While Not orstemp1.EOF
        With mrs_Parametros
            .AddNew
            .Fields!IdParametro = orstemp1!IdParametro
            .Fields!tipo = orstemp1!tipo
            .Fields!Codigo = orstemp1!Codigo
            .Fields!ValorTexto = orstemp1!ValorTexto
            .Fields!ValorInt = orstemp1!ValorInt
            .Fields!ValorFloat = orstemp1!ValorFloat
            .Fields!descripcion = orstemp1!descripcion
            .Fields!Grupo = orstemp1!Grupo
            .Update
        End With
        orstemp1.MoveNext
        Loop
        mrs_Parametros.MoveFirst
    End If
    
End Sub

Private Sub Form_Initialize()
    Set mo_cmbGrupo.MiComboBox = cmbGrupo
End Sub

Private Sub btnAceptar_Click()
    If ValidarDatosObligatorios() Then
        If ValidarReglas() Then
            If ModificarDatos() Then
                MsgBox "                  Los datos se modificaron correctamente                   " & Chr(13) & _
                       "para que haga efecto el cambio, debe salir del Sistema y volver a ingresar", vbInformation, Me.Caption
                Me.Visible = False
            Else
                MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasComunes.MensajeError, vbExclamation, Me.Caption
            End If
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Private Sub Form_Load()
    Dim rsDocumentos As New Recordset
    GenerarRecordsetTemporal
    If ml_Grupo <> "" Then
    Set rsDocumentos = mo_ReglasComunes.SeleccionarGrupoParametros(ml_Grupo)
    Do While Not rsDocumentos.EOF
        With mrs_Parametros
            .AddNew
            .Fields!IdParametro = rsDocumentos!IdParametro
            .Fields!tipo = rsDocumentos!tipo
            .Fields!Codigo = rsDocumentos!Codigo
            .Fields!ValorTexto = rsDocumentos!ValorTexto
            .Fields!ValorInt = rsDocumentos!ValorInt
            .Fields!ValorFloat = rsDocumentos!ValorFloat
            .Fields!descripcion = rsDocumentos!descripcion
            .Fields!Grupo = rsDocumentos!Grupo
        End With
        rsDocumentos.MoveNext
    Loop
    End If
    mo_Apariencia.ConfigurarFilasBiColores Me.grdParametros, sighEntidades.GrillaConFilasBicolor
    mo_cmbGrupo.BoundColumn = "IdParametro"
    mo_cmbGrupo.ListField = "Grupo"
    Set mo_cmbGrupo.RowSource = mo_ReglasComunes.LlenadoParametros()
End Sub

Private Sub grdParametros_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Not (Cell.Column.Key = "ValorTexto" Or Cell.Column.Key = "ValorInt") Then
        Cancel = True
    End If
End Sub

Private Sub grdParametros_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    grdParametros.Override.DefaultRowHeight = 600
    
    grdParametros.Bands(0).Columns("IdParametro").Width = 800
    grdParametros.Bands(0).Columns("IdParametro").Activation = ssActivationActivateNoEdit
    
    grdParametros.Bands(0).Columns("Tipo").Hidden = True
    grdParametros.Bands(0).Columns("Codigo").Hidden = True

    grdParametros.Bands(0).Columns("ValorTexto").Width = 1000
    grdParametros.Bands(0).Columns("ValorTexto").Header.Caption = "Texto"
    grdParametros.Bands(0).Columns("ValorTexto").Header.Appearance.BackColor = vbRed
    
    grdParametros.Bands(0).Columns("ValorInt").Width = 1000
    grdParametros.Bands(0).Columns("ValorInt").Header.Caption = "Entero"
    grdParametros.Bands(0).Columns("ValorInt").Header.Appearance.BackColor = vbRed
    
    grdParametros.Bands(0).Columns("ValorFloat").Hidden = True
    grdParametros.Bands(0).Columns("Descripcion").Width = 7000
    grdParametros.Bands(0).Columns("Descripcion").Header.Caption = "Descripción"
    grdParametros.Bands(0).Columns("Descripcion").Activation = ssActivationActivateNoEdit
    grdParametros.Bands(0).Columns("Descripcion").CellMultiLine = ssCellMultiLineTrue
    
    grdParametros.Bands(0).Columns("Grupo").Width = 1500
    grdParametros.Bands(0).Columns("Grupo").Activation = ssActivationActivateNoEdit
    
    mo_Apariencia.ConfigurarFilasBiColores Me.grdParametros, sighEntidades.GrillaConFilasBicolor
End Sub

Sub GenerarRecordsetTemporal()
    Set mrs_Parametros = New ADODB.Recordset
    With mrs_Parametros
          .Fields.Append "IdParametro", adInteger, 4, adFldIsNullable
          .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "ValorTexto", adVarChar, 355, adFldIsNullable
          .Fields.Append "ValorInt", adInteger, 4, adFldIsNullable
          .Fields.Append "ValorFloat", adDouble, 8, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 150, adFldIsNullable
          .Fields.Append "Grupo", adVarChar, 30, adFldIsNullable
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdParametros.DataSource = mrs_Parametros
    mo_Apariencia.ConfigurarFilasBiColores Me.grdParametros, sighEntidades.GrillaConFilasBicolor
End Sub

Function ModificarDatos() As Boolean
Dim oConexion As New ADODB.Connection
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighEntidades.CadenaConexion
Set oParametros.Conexion = oConexion
    With mo_Param
        If mrs_Parametros.RecordCount > 0 Then
            mrs_Parametros.MoveFirst
            Do While Not mrs_Parametros.EOF
                .IdParametro = mrs_Parametros.Fields!IdParametro
                .tipo = mrs_Parametros.Fields!tipo
                .Codigo = IIf(IsNull(mrs_Parametros.Fields!Codigo), "", mrs_Parametros.Fields!Codigo)
                If IsNull(mrs_Parametros.Fields!ValorTexto) Then
                    .ValorTexto = ""
                Else
                    .ValorTexto = mrs_Parametros.Fields!ValorTexto
                End If
                
                If IsNull(mrs_Parametros.Fields!ValorInt) Then
                    .ValorInt = 0
                Else
                    .ValorInt = mrs_Parametros.Fields!ValorInt
                End If
                
                If IsNull(mrs_Parametros.Fields!ValorFloat) Then
                    .ValorFloat = 0
                Else
                    .ValorFloat = mrs_Parametros.Fields!ValorFloat
                End If
                .descripcion = mrs_Parametros!descripcion
                .Grupo = mrs_Parametros.Fields!Grupo
                
                If Not oParametros.Modificar(mo_Param) Then
                    ModificarDatos = False
                    Exit Function
                End If
                mrs_Parametros.MoveNext
            Loop
        End If
    End With
    ModificarDatos = True
    oConexion.Close
    Set oConexion = Nothing
End Function

Function ValidarDatosObligatorios() As Boolean
    Dim sMensaje As String
    ValidarDatosObligatorios = False
    
    If Not (mrs_Parametros.EOF = True And mrs_Parametros.BOF = True) Then
         mrs_Parametros.MoveFirst
    End If
   
    ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
   ValidarReglas = False
   ValidarReglas = True
End Function
