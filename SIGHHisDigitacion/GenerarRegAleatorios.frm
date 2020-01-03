VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmGenerarRegAleatorios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HIS - Generar registros aleatorios para la doble digitación"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "GenerarRegAleatorios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chbAlAzar 
      Caption         =   "Al azar"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   7440
      MaxLength       =   3
      TabIndex        =   18
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtTotalRegistros 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8115
      Begin VB.ComboBox cmbEstablecimiento 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
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
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskfechaAnio 
         Height          =   330
         Left            =   7320
         TabIndex        =   3
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Establecimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   8055
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "GenerarRegAleatorios.frx":000C
         DownPicture     =   "GenerarRegAleatorios.frx":046C
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
         Left            =   2640
         Picture         =   "GenerarRegAleatorios.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "GenerarRegAleatorios.frx":0D56
         DownPicture     =   "GenerarRegAleatorios.frx":121A
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
         Left            =   4200
         Picture         =   "GenerarRegAleatorios.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvResumenHIS 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   7011
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ugvResumenHIS"
   End
   Begin UltraGrid.SSUltraGrid grdRegAleatorios 
      Height          =   3975
      Left            =   4080
      TabIndex        =   13
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7011
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "grdRegAleatorios"
   End
   Begin Threed.SSCommand btnGenerar 
      Height          =   465
      Left            =   1080
      TabIndex        =   14
      Top             =   5520
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   820
      _Version        =   262144
      PictureFrames   =   1
      Picture         =   "GenerarRegAleatorios.frx":1BF2
      Caption         =   "Generar registros aleatorios para la doble digitación al 95% de confianza  "
      PictureAlignment=   9
   End
   Begin VB.Label lblRegistrosMuestra 
      Caption         =   "Total de Registros al 95% (Muestra)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Total de Registros del Lote (Población)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "frmGenerarRegAleatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Generar registors aleatorios
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_DatosParametros As New SIGHDatos.Parametros
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_cmbEstablecimiento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim oRcsRegistrosLote As New ADODB.Recordset
Dim oRcsRegLoteMuestra As New ADODB.Recordset
Dim oTablaDOHIS_Lote As New DOHIS_Lotes
Dim mi_Opcion As sghOpciones
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_IdEstablecimiento As Long
Dim ml_NroTotalHojas As Integer
Dim ml_HojasRegistradas As Integer
Dim ml_IdUsuario As Long
Dim ml_IdLote As Long
Dim ml_Muestra As Long
Dim ml_Confianza As Integer
Dim ms_fechaactual As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim Vector() As Variant

'========================================== PROPIEDADES ========================================
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property
Property Get IdLote() As Long
    IdLote = ml_IdLote
End Property
Property Let IdLote(lValue As Long)
   ml_IdLote = lValue
End Property

Sub GenerarRegistrosAleatorios()
'    Dim lnRegAleatorios(1 To 1500) As Integer
'    Dim lnRegOrdenados(ml_Muestra) As Integer
    Dim lnValorMax As Integer
    Dim N, M, Al As Integer
    Dim Repite As Boolean
    'Genera el vector de registros
    
    ReDim Vector(ml_Muestra - 1)
    
    If chbAlAzar.Value = 1 Then
       Vector(0) = Int((ml_Muestra - 1) * Rnd + 1)
       For N = 1 To ml_Muestra - 1
          Randomize
          Do
             Repite = False
             Al = Redondeo((ml_Muestra - 1) * Rnd + 1, 0)
             For M = 0 To N - 1
                If Al = Vector(M) Then
                   Repite = True
                End If
             Next
           Loop While Repite = True
           Vector(N) = Al
       Next
    Else
       For N = 0 To ml_Muestra - 1
          Vector(N) = N + 1
       Next
    End If
    'Ordenar
'    Dim i As Integer, j As Integer, temp As Integer
'    For i = 1 To UBound(lnRegAleatorios) - 1
'        For j = i To UBound(lnRegAleatorios)
'            If lnRegAleatorios(i) > lnRegAleatorios(j) Then
'                temp = lnRegAleatorios(i)
'                lnRegAleatorios(i) = lnRegAleatorios(j)
'                lnRegAleatorios(j) = temp
'            End If
'        Next
'    Next

    Ordenar_Matriz Vector, 0, ml_Muestra - 1
    
    'Limpia muestra
    With oRcsRegLoteMuestra
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
    
    'Cargar en la tabla muestra
    For N = 0 To ml_Muestra - 1
        If oRcsRegistrosLote.RecordCount > 0 Then
           oRcsRegistrosLote.MoveFirst
           Do While Not oRcsRegistrosLote.EOF
              If Vector(N) = oRcsRegistrosLote.Fields!NroRegistroLote Then
                    oRcsRegLoteMuestra.AddNew
                    oRcsRegLoteMuestra.Fields!NroRegistroLote = oRcsRegistrosLote.Fields!NroRegistroLote
                    oRcsRegLoteMuestra.Fields!IdHisLote = oRcsRegistrosLote.Fields!IdHisLote
                    oRcsRegLoteMuestra.Fields!Lote = oRcsRegistrosLote.Fields!Lote
                    oRcsRegLoteMuestra.Fields!IdHisCabecera = oRcsRegistrosLote.Fields!IdHisCabecera
                    oRcsRegLoteMuestra.Fields!IdEstablecimiento = oRcsRegistrosLote.Fields!IdEstablecimiento
                    oRcsRegLoteMuestra.Fields!NroHojaHis = oRcsRegistrosLote.Fields!NroHojaHis
                    oRcsRegLoteMuestra.Fields!IdHisDetalle = oRcsRegistrosLote.Fields!IdHisDetalle
                    oRcsRegLoteMuestra.Fields!NroRegistroHoja = oRcsRegistrosLote.Fields!NroRegistroHoja
                    oRcsRegLoteMuestra.Fields!DiaAtencion = oRcsRegistrosLote.Fields!DiaAtencion
                    oRcsRegLoteMuestra.Fields!IdTipoAtencion = oRcsRegistrosLote.Fields!IdTipoAtencion
                    oRcsRegLoteMuestra.Fields!HC_FF_COD = oRcsRegistrosLote.Fields!HC_FF_COD
                    oRcsRegLoteMuestra.Fields!IdPais = oRcsRegistrosLote.Fields!IdPais
                    oRcsRegLoteMuestra.Fields!Codigo = oRcsRegistrosLote.Fields!Codigo
                    oRcsRegLoteMuestra.Fields!IdTipoDocumento = oRcsRegistrosLote.Fields!IdTipoDocumento
                    oRcsRegLoteMuestra.Fields!Documento = oRcsRegistrosLote.Fields!Documento
                    oRcsRegLoteMuestra.Fields!NroDocIdentidad = oRcsRegistrosLote.Fields!NroDocIdentidad
                    oRcsRegLoteMuestra.Fields!NroHijo = oRcsRegistrosLote.Fields!NroHijo
                    oRcsRegLoteMuestra.Fields!IdTipoFinanciamiento = oRcsRegistrosLote.Fields!IdTipoFinanciamiento
                    oRcsRegLoteMuestra.Fields!Financiamiento = oRcsRegistrosLote.Fields!Financiamiento
                    oRcsRegLoteMuestra.Fields!IdEtnia = oRcsRegistrosLote.Fields!IdEtnia
                    oRcsRegLoteMuestra.Fields!Etnia = oRcsRegistrosLote.Fields!Etnia
                    oRcsRegLoteMuestra.Fields!IdDistrito = oRcsRegistrosLote.Fields!IdDistrito
                    oRcsRegLoteMuestra.Fields!Distrito = oRcsRegistrosLote.Fields!Distrito
                    oRcsRegLoteMuestra.Fields!Edad = oRcsRegistrosLote.Fields!Edad
                    oRcsRegLoteMuestra.Fields!IdTipoEdad = oRcsRegistrosLote.Fields!IdTipoEdad
                    oRcsRegLoteMuestra.Fields!TipoEdad = oRcsRegistrosLote.Fields!TipoEdad
                    oRcsRegLoteMuestra.Fields!Sexo = oRcsRegistrosLote.Fields!Sexo
                    oRcsRegLoteMuestra.Fields!Peso = oRcsRegistrosLote.Fields!Peso
                    oRcsRegLoteMuestra.Fields!Talla = oRcsRegistrosLote.Fields!Talla
                    oRcsRegLoteMuestra.Fields!IdEstadoaEstablec = oRcsRegistrosLote.Fields!IdEstadoaEstablec
                    oRcsRegLoteMuestra.Fields!IdEstadoaServicio = oRcsRegistrosLote.Fields!IdEstadoaServicio
                    oRcsRegLoteMuestra.Update
              End If
              oRcsRegistrosLote.MoveNext
           Loop
           oRcsRegistrosLote.MoveFirst
           If oRcsRegLoteMuestra.RecordCount > 0 Then oRcsRegLoteMuestra.MoveFirst
        End If
    Next
End Sub

Sub Ordenar_Matriz(El_Vector() As Variant, _
                   Limite_Inferior As Long, _
                   Limite_Superior As Long)
  
    Dim i As Long, j As Long, x As Variant, y As Variant
      
    i = Limite_Inferior
    j = Limite_Superior
      
    x = El_Vector((Limite_Inferior + Limite_Superior) / 2)
      
    While i <= j
          
        While (El_Vector(i) < x) And (i < Limite_Superior)
            i = i + 1
        Wend
          
        While (x < El_Vector(j)) And (j > Limite_Inferior)
            j = j - 1
        Wend
          
        If i <= j Then
            y = El_Vector(i)
            El_Vector(i) = El_Vector(j)
            El_Vector(j) = y
            i = i + 1
            j = j - 1
        End If
      
    Wend
      
    If Limite_Inferior < j Then Ordenar_Matriz El_Vector(), Limite_Inferior, j
    If i < Limite_Superior Then Ordenar_Matriz El_Vector(), i, Limite_Superior
  
End Sub

Private Sub btnGenerar_Click()
    GenerarRegistrosAleatorios
End Sub




Private Sub chbAlAzar_KeyDown(KeyCode As Integer, Shift As Integer)
AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub cmbEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
AdministrarKeyPreview CInt(KeyCode)
End Sub

'========================================== EVENTOS ========================================
Private Sub Form_Load()
    Dim oRcsTemp As ADODB.Recordset
    Set mo_cmbMes.MiComboBox = Me.cmbMes
    Set mo_cmbEstablecimiento.MiComboBox = Me.cmbEstablecimiento
    
    mo_Formulario.HabilitarDeshabilitar Me.cmbEstablecimiento, False
    mo_Formulario.HabilitarDeshabilitar Me.txtLote, False
    mo_Formulario.HabilitarDeshabilitar Me.txtTotalRegistros, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbMes, False
    mo_Formulario.HabilitarDeshabilitar Me.mskfechaAnio, False
    mo_Formulario.HabilitarDeshabilitar Me.Text1, False
    
    CargarComboBoxes
    CreaTemporaloRsRegistrosLote
    CargarDatosAlFormulario
End Sub

Sub CreaTemporaloRsRegistrosLote()
    If oRcsRegistrosLote.State = 1 Then
       Set oRcsRegistrosLote = Nothing
    End If
    With oRcsRegistrosLote
          .Fields.Append "NroRegistroLote", adInteger
          .Fields.Append "IdHisLote", adInteger
          .Fields.Append "Lote", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdHisCabecera", adInteger, adFldIsNullable
          .Fields.Append "IdEstablecimiento", adInteger, adFldIsNullable
          .Fields.Append "NroHojaHis", adInteger, adFldIsNullable
          .Fields.Append "IdHisDetalle", adInteger, adFldIsNullable
          .Fields.Append "NroRegistroHoja", adInteger, adFldIsNullable
          .Fields.Append "DiaAtencion", adInteger, adFldIsNullable
          .Fields.Append "IdTipoAtencion", adInteger, adFldIsNullable
          .Fields.Append "HC_FF_COD", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdPais", adInteger, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdTipoDocumento", adInteger, adFldIsNullable
          .Fields.Append "Documento", adVarChar, 255, adFldIsNullable
          .Fields.Append "NroDocIdentidad", adVarChar, 255, adFldIsNullable
          .Fields.Append "NroHijo", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdTipoFinanciamiento", adInteger, adFldIsNullable
          .Fields.Append "Financiamiento", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdEtnia", adInteger, adFldIsNullable
          .Fields.Append "Etnia", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdDistrito", adInteger, adFldIsNullable
          .Fields.Append "Distrito", adVarChar, 255, adFldIsNullable
          .Fields.Append "Edad", adInteger, adFldIsNullable
          .Fields.Append "IdTipoEdad", adInteger, adFldIsNullable
          .Fields.Append "TipoEdad", adVarChar, 255, adFldIsNullable
          .Fields.Append "Sexo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Peso", adVarChar, 255, adFldIsNullable
          .Fields.Append "Talla", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdEstadoaEstablec", adInteger, adFldIsNullable
          .Fields.Append "IdEstadoaServicio", adInteger, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    
    If oRcsRegLoteMuestra.State = 1 Then
       Set oRcsRegLoteMuestra = Nothing
    End If
    With oRcsRegLoteMuestra
          .Fields.Append "NroRegistroLote", adInteger
          .Fields.Append "IdHisLote", adInteger
          .Fields.Append "Lote", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdHisCabecera", adInteger, adFldIsNullable
          .Fields.Append "IdEstablecimiento", adInteger, adFldIsNullable
          .Fields.Append "NroHojaHis", adInteger, adFldIsNullable
          .Fields.Append "IdHisDetalle", adInteger, adFldIsNullable
          .Fields.Append "NroRegistroHoja", adInteger, adFldIsNullable
          .Fields.Append "DiaAtencion", adInteger, adFldIsNullable
          .Fields.Append "IdTipoAtencion", adInteger, adFldIsNullable
          .Fields.Append "HC_FF_COD", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdPais", adInteger, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdTipoDocumento", adInteger, adFldIsNullable
          .Fields.Append "Documento", adVarChar, 255, adFldIsNullable
          .Fields.Append "NroDocIdentidad", adVarChar, 255, adFldIsNullable
          .Fields.Append "NroHijo", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdTipoFinanciamiento", adInteger, adFldIsNullable
          .Fields.Append "Financiamiento", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdEtnia", adInteger, adFldIsNullable
          .Fields.Append "Etnia", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdDistrito", adInteger, adFldIsNullable
          .Fields.Append "Distrito", adVarChar, 255, adFldIsNullable
          .Fields.Append "Edad", adInteger, adFldIsNullable
          .Fields.Append "IdTipoEdad", adInteger, adFldIsNullable
          .Fields.Append "TipoEdad", adVarChar, 255, adFldIsNullable
          .Fields.Append "Sexo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Peso", adVarChar, 255, adFldIsNullable
          .Fields.Append "Talla", adVarChar, 255, adFldIsNullable
          .Fields.Append "IdEstadoaEstablec", adInteger, adFldIsNullable
          .Fields.Append "IdEstadoaServicio", adInteger, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdRegAleatorios.DataSource = oRcsRegLoteMuestra
    mo_Apariencia.ConfigurarFilasBiColores Me.grdRegAleatorios, SIGHEntidades.GrillaConFilasBicolor
    
End Sub

Private Sub btnAceptar_Click()
    If ValidarDatosObligatorios Then
        If ValidarReglas Then
            If IngresarDatos Then
                Call MsgBox("Se guardo correctamente los registros generados para la doble digitación", vbInformation, Me.Caption)
                Me.Hide
            Else
                Call MsgBox("No se pudo guardar los registros generados, Verificar Error.", vbExclamation, Me.Caption)
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub btnCancelar_Click()
    Me.Hide
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtTotalRegistros_KeyDown(KeyCode As Integer, Shift As Integer)
AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub txtTotalRegistros_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
        KeyAscii = 8
    Else
        KeyAscii = 1
    End If
End If
End Sub

Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbMes
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub cmbMes_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, mskfechaAnio
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub mskfechaAnio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, btnAceptar
    AdministrarKeyPreview CInt(KeyCode)
End Sub

'========================================== METODOS ========================================
Sub CargarDatosAlFormulario()
    Dim oRcsTemp As New ADODB.Recordset
    oTablaDOHIS_Lote.IdHisLote = ml_IdLote
    Set oTablaDOHIS_Lote = mo_ReglasHIS.ConsultarRegistroLoteHIS(oTablaDOHIS_Lote)
    ms_fechaactual = mo_DatosParametros.RetornaFechaServidorSQL
    mo_cmbEstablecimiento.BoundText = ml_IdEstablecimiento
    Me.txtLote.Text = oTablaDOHIS_Lote.Lote
    mo_cmbMes.BoundText = oTablaDOHIS_Lote.Mes
    Me.mskfechaAnio.Text = oTablaDOHIS_Lote.Anio
    Set oRcsTemp = mo_ReglasHIS.His_ConsultarTotalRegistrosLote(ml_IdEstablecimiento, ml_IdLote)
    Me.txtTotalRegistros.Text = oRcsTemp.RecordCount
    
    ml_Muestra = 0
    ml_Confianza = 0
    Select Case Val(lcBuscaParametro.SeleccionaFilaParametro(331))
    Case 1
        ml_Confianza = 95
    Case 2
        ml_Confianza = 99
    End Select
    ml_Muestra = CalculaMuestra(Val(Me.txtTotalRegistros.Text), ml_Confianza)
    Me.Text1.Text = ml_Muestra
    Me.lblRegistrosMuestra.Caption = "Total de Registros al " & CStr(ml_Confianza) & "% (Muestra)"
    Me.btnGenerar.Caption = "Generar registros aleatorios para la doble digitación al " & CStr(ml_Confianza) & "% de confianza  (F7)"
    
    'CARGAR GRILLA
    Set oRcsTemp = mo_ReglasHIS.HIS_ConsultarRegistrosTotalesLotes(ml_IdLote)
    If oRcsTemp.RecordCount > 0 Then
       oRcsTemp.MoveFirst
       Do While Not oRcsTemp.EOF
          oRcsRegistrosLote.AddNew
          oRcsRegistrosLote.Fields!NroRegistroLote = oRcsTemp.Fields!NroRegistroLote
          oRcsRegistrosLote.Fields!IdHisLote = oRcsTemp.Fields!IdHisLote
          oRcsRegistrosLote.Fields!Lote = oRcsTemp.Fields!Lote
          oRcsRegistrosLote.Fields!IdHisCabecera = oRcsTemp.Fields!IdHisCabecera
          oRcsRegistrosLote.Fields!IdEstablecimiento = oRcsTemp.Fields!IdEstablecimiento
          oRcsRegistrosLote.Fields!NroHojaHis = oRcsTemp.Fields!NroHojaHis
          oRcsRegistrosLote.Fields!IdHisDetalle = oRcsTemp.Fields!IdHisDetalle
          oRcsRegistrosLote.Fields!NroRegistroHoja = oRcsTemp.Fields!NroRegistroHoja
          oRcsRegistrosLote.Fields!DiaAtencion = oRcsTemp.Fields!DiaAtencion
          oRcsRegistrosLote.Fields!IdTipoAtencion = oRcsTemp.Fields!IdTipoAtencion
          oRcsRegistrosLote.Fields!HC_FF_COD = oRcsTemp.Fields!HC_FF_COD
          oRcsRegistrosLote.Fields!IdPais = IIf(IsNull(oRcsTemp.Fields!IdPais), 0, oRcsTemp.Fields!IdPais)
          oRcsRegistrosLote.Fields!Codigo = oRcsTemp.Fields!Codigo
          oRcsRegistrosLote.Fields!IdTipoDocumento = IIf(IsNull(oRcsTemp.Fields!IdTipoDocumento), 0, oRcsTemp.Fields!IdTipoDocumento)
          oRcsRegistrosLote.Fields!Documento = oRcsTemp.Fields!Documento
          oRcsRegistrosLote.Fields!NroDocIdentidad = oRcsTemp.Fields!NroDocIdentidad
          oRcsRegistrosLote.Fields!NroHijo = oRcsTemp.Fields!NroHijo
          oRcsRegistrosLote.Fields!IdTipoFinanciamiento = IIf(IsNull(oRcsTemp.Fields!IdTipoFinanciamiento), 0, oRcsTemp.Fields!IdTipoFinanciamiento)
          oRcsRegistrosLote.Fields!Financiamiento = oRcsTemp.Fields!Financiamiento
          oRcsRegistrosLote.Fields!IdEtnia = IIf(IsNull(oRcsTemp.Fields!IdEtnia), 0, oRcsTemp.Fields!IdEtnia)
          oRcsRegistrosLote.Fields!Etnia = oRcsTemp.Fields!Etnia
          oRcsRegistrosLote.Fields!IdDistrito = IIf(IsNull(oRcsTemp.Fields!IdDistrito), 0, oRcsTemp.Fields!IdDistrito)
          oRcsRegistrosLote.Fields!Distrito = oRcsTemp.Fields!Distrito
          oRcsRegistrosLote.Fields!Edad = IIf(IsNull(oRcsTemp.Fields!Edad), 0, oRcsTemp.Fields!Edad)
          oRcsRegistrosLote.Fields!IdTipoEdad = IIf(IsNull(oRcsTemp.Fields!IdTipoEdad), 0, oRcsTemp.Fields!IdTipoEdad)
          oRcsRegistrosLote.Fields!TipoEdad = oRcsTemp.Fields!TipoEdad
          oRcsRegistrosLote.Fields!Sexo = oRcsTemp.Fields!Sexo
          oRcsRegistrosLote.Fields!Peso = oRcsTemp.Fields!Peso
          oRcsRegistrosLote.Fields!Talla = oRcsTemp.Fields!Talla
          oRcsRegistrosLote.Fields!IdEstadoaEstablec = IIf(IsNull(oRcsTemp.Fields!IdEstadoaEstablec), 0, oRcsTemp.Fields!IdEstadoaEstablec)
          oRcsRegistrosLote.Fields!IdEstadoaServicio = IIf(IsNull(oRcsTemp.Fields!IdEstadoaServicio), 0, oRcsTemp.Fields!IdEstadoaServicio)
          oRcsRegistrosLote.Update
          oRcsTemp.MoveNext
       Loop
       oRcsRegistrosLote.MoveFirst
    End If
    Set Me.ugvResumenHIS.DataSource = oRcsRegistrosLote
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvResumenHIS, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Function Redondeo(ByVal Numero, ByVal Decimales)
      Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function

Sub CargarComboBoxes()
    mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
    mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
    Set mo_cmbEstablecimiento.RowSource = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    mo_cmbEstablecimiento.BoundText = ml_IdEstablecimiento
   
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
End Sub

Private Function ListadoEstablecimientos() As Recordset
    Dim oTabla As New DOEstablecimiento
    Dim oRcs_Establecimiento As New Recordset
    Set ListadoEstablecimientos = mo_ReglasHIS.ObtenerListaEstablecimientosMR
End Function

Function ValidarDatosObligatorios() As Boolean
    On Error Resume Next
    ValidarDatosObligatorios = False
    If oRcsRegLoteMuestra.RecordCount = 0 Then
        Call MsgBox("Debe generar previamente los registros aleatorios", vbInformation, Me.Caption)
        Exit Function
    End If
    ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
    ValidarReglas = False
'    CargaDatosAlObjetosDeDatos
'    If mo_ReglasHIS.ValidarLoteHIS_LoteExiste(oTablaDOHIS_Lote) Then
'        Call MsgBox("El código del lote Existe, elija otro código.", vbCritical, Me.Caption)
'        Exit Function
'    End If
    ValidarReglas = True
End Function

Function IngresarDatos() As Boolean
    Dim IngresarRegAleatorios As Boolean
    Dim mbActualizoRegistro As Boolean
    Dim oRcsTemp As New ADODB.Recordset
    Dim lnNroRegistro As Long
    IngresarRegAleatorios = mo_ReglasHIS.IngresarHISVerificado(oRcsRegLoteMuestra)
    If IngresarRegAleatorios Then
        oTablaDOHIS_Lote.DobleDigitacion = 1
        IngresarDatos = mo_ReglasHIS.ModificarRegistroLoteHIS(oTablaDOHIS_Lote)
    Else
        IngresarDatos = False
    End If
End Function

Sub CargaDatosAlObjetosDeDatos()
End Sub

Sub LimpiarVariablesDeMemoria()
Set mo_ReglasHIS = Nothing
Set mo_cmbMes = Nothing
End Sub

Private Sub txtNroPag_KeyPress(KeyAscii As Integer)
If ((KeyAscii < 48) Or KeyAscii > 57) Then
    If KeyAscii = 8 Then
        KeyAscii = 8
    Else
        If KeyAscii = 46 Then
            KeyAscii = 46
        Else
            KeyAscii = 1
        End If
    End If
End If
End Sub

'Calcula número de registros aleatorios
Function CalculaMuestra(lnPoblacion As Integer, lnConfianza As Integer) As Long
    CalculaMuestra = 0
    If lnPoblacion > 0 And lnConfianza > 0 Then
       On Error Resume Next
       Dim EXL As Excel.Application
       Set EXL = New Excel.Application
       Dim W As Excel.Workbook
       Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\formula_muestras_finitas.xls")

       Dim s As Excel.Worksheet
       Set s = W.Sheets("Datos")
       s.Cells(2, 2).Value = lnPoblacion
       s.Cells(2, 6).Value = lnConfianza
       CalculaMuestra = Redondeo(s.Cells(6, 2).Value, 0)

       W.Close False
       Set s = Nothing
       Set W = Nothing
       Set EXL = Nothing
    End If
End Function

Private Sub ugvResumenHIS_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With Me.ugvResumenHIS.Bands(0)
        .Columns("NroRegistroLote").Header.Caption = "ID"
        .Columns("NroRegistroLote").Width = 400
        .Columns("NroRegistroLote").Activation = ssActivationActivateNoEdit
        .Columns("IdHisLote").Hidden = True
        .Columns("Lote").Header.Caption = "Lote"
        .Columns("Lote").Width = 600
        .Columns("Lote").Activation = ssActivationActivateNoEdit
        .Columns("IdHisCabecera").Hidden = True
        .Columns("IdEstablecimiento").Hidden = True
        .Columns("NroHojaHis").Header.Caption = "Nro Hoja"
        .Columns("NroHojaHis").Width = 700
        .Columns("NroHojaHis").Activation = ssActivationActivateNoEdit
        .Columns("IdHisDetalle").Hidden = True
        .Columns("NroRegistroHoja").Header.Caption = "Registro Hoja"
        .Columns("NroRegistroHoja").Width = 1100
        .Columns("NroRegistroHoja").Activation = ssActivationActivateNoEdit
        .Columns("DiaAtencion").Header.Caption = "Día"
        .Columns("DiaAtencion").Width = 500
        .Columns("DiaAtencion").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoAtencion").Hidden = True
        .Columns("HC_FF_COD").Header.Caption = "HC_FF_COD"
        .Columns("HC_FF_COD").Width = 1100
        .Columns("HC_FF_COD").Activation = ssActivationActivateNoEdit
        .Columns("IdPais").Hidden = True
        .Columns("Codigo").Header.Caption = "País"
        .Columns("Codigo").Width = 500
        .Columns("Codigo").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoDocumento").Hidden = True
        .Columns("Documento").Header.Caption = "Tipo Doc."
        .Columns("Documento").Width = 1000
        .Columns("Documento").Activation = ssActivationActivateNoEdit
        .Columns("NroDocIdentidad").Header.Caption = "Nro Doc. Ident"
        .Columns("NroDocIdentidad").Width = 1100
        .Columns("NroDocIdentidad").Activation = ssActivationActivateNoEdit
        .Columns("NroHijo").Header.Caption = "Nro Hijo"
        .Columns("NroHijo").Width = 650
        .Columns("NroHijo").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoFinanciamiento").Hidden = True
        .Columns("Financiamiento").Hidden = True
        .Columns("IdEtnia").Hidden = True
        .Columns("Etnia").Hidden = True
        .Columns("IdDistrito").Hidden = True
        .Columns("Distrito").Hidden = True
        .Columns("Edad").Hidden = True
        .Columns("IdTipoEdad").Hidden = True
        .Columns("TipoEdad").Hidden = True
        .Columns("Sexo").Hidden = True
        .Columns("Peso").Hidden = True
        .Columns("Talla").Hidden = True
        .Columns("IdEstadoaEstablec").Hidden = True
        .Columns("IdEstadoaServicio").Hidden = True
    End With
   
End Sub


Private Sub grdRegAleatorios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With grdRegAleatorios.Bands(0)
        .Columns("NroRegistroLote").Header.Caption = "ID"
        .Columns("NroRegistroLote").Width = 400
        .Columns("NroRegistroLote").Activation = ssActivationActivateNoEdit
        .Columns("IdHisLote").Hidden = True
        .Columns("Lote").Header.Caption = "Lote"
        .Columns("Lote").Width = 600
        .Columns("Lote").Activation = ssActivationActivateNoEdit
        .Columns("IdHisCabecera").Hidden = True
        .Columns("IdEstablecimiento").Hidden = True
        .Columns("NroHojaHis").Header.Caption = "Nro Hoja"
        .Columns("NroHojaHis").Width = 700
        .Columns("NroHojaHis").Activation = ssActivationActivateNoEdit
        .Columns("IdHisDetalle").Hidden = True
        .Columns("NroRegistroHoja").Header.Caption = "Registro Hoja"
        .Columns("NroRegistroHoja").Width = 1100
        .Columns("NroRegistroHoja").Activation = ssActivationActivateNoEdit
        .Columns("DiaAtencion").Header.Caption = "Día"
        .Columns("DiaAtencion").Width = 500
        .Columns("DiaAtencion").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoAtencion").Hidden = True
        .Columns("HC_FF_COD").Header.Caption = "HC_FF_COD"
        .Columns("HC_FF_COD").Width = 1100
        .Columns("HC_FF_COD").Activation = ssActivationActivateNoEdit
        .Columns("IdPais").Hidden = True
        .Columns("Codigo").Header.Caption = "País"
        .Columns("Codigo").Width = 500
        .Columns("Codigo").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoDocumento").Hidden = True
        .Columns("Documento").Header.Caption = "Tipo Doc."
        .Columns("Documento").Width = 1000
        .Columns("Documento").Activation = ssActivationActivateNoEdit
        .Columns("NroDocIdentidad").Header.Caption = "Nro Doc. Ident"
        .Columns("NroDocIdentidad").Width = 1100
        .Columns("NroDocIdentidad").Activation = ssActivationActivateNoEdit
        .Columns("NroHijo").Header.Caption = "Nro Hijo"
        .Columns("NroHijo").Width = 650
        .Columns("NroHijo").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoFinanciamiento").Hidden = True
        .Columns("Financiamiento").Hidden = True
        .Columns("IdEtnia").Hidden = True
        .Columns("Etnia").Hidden = True
        .Columns("IdDistrito").Hidden = True
        .Columns("Distrito").Hidden = True
        .Columns("Edad").Hidden = True
        .Columns("IdTipoEdad").Hidden = True
        .Columns("TipoEdad").Hidden = True
        .Columns("Sexo").Hidden = True
        .Columns("Peso").Hidden = True
        .Columns("Talla").Hidden = True
        .Columns("IdEstadoaEstablec").Hidden = True
        .Columns("IdEstadoaServicio").Hidden = True
    End With
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
     Case vbKeyF7
        btnGenerar_Click
     Case vbKeyF8
    End Select
End Sub

Private Sub ugvResumenHIS_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    AdministrarKeyPreview CInt(KeyCode)
End Sub
