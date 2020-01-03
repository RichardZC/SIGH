VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcCEinmunizaciones 
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   ScaleHeight     =   3840
   ScaleWidth      =   6735
   Begin UltraGrid.SSUltraGrid grdBienes 
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6641
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   68157460
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Override        =   "UcCEinmunizaciones.ctx":0000
      Caption         =   "Lista de Inmunizaciones"
   End
End
Attribute VB_Name = "UcCEinmunizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Inicializar()

    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbIdServicio.MiComboBox = cmbIdServicio

    ConfigurarTipoServicio
    ConfiguraPermisos
    cmbIdTipoServicio_Click
End Function

