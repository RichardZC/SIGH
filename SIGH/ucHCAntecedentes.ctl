VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.UserControl ucHCAntecedentes 
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   LockControls    =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   10515
   Begin TabDlg.SSTab sTabAntecedentes 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Antec. Perinatales"
      TabPicture(0)   =   "ucHCAntecedentes.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Nacimiento"
      TabPicture(1)   =   "ucHCAntecedentes.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame(3)"
      Tab(1).Control(1)=   "Frame(2)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Alimentación/Patológicos"
      TabPicture(2)   =   "ucHCAntecedentes.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame(1)"
      Tab(2).Control(1)=   "Frame(0)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Ant. Fam/Vivienda"
      TabPicture(3)   =   "ucHCAntecedentes.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame(5)"
      Tab(3).Control(1)=   "Frame(4)"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame 
         Caption         =   "2. Alimentaciòn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   1
         Left            =   -74880
         TabIndex        =   144
         Top             =   480
         Width           =   5055
         Begin VB.Frame frAlimentacion 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   2160
            TabIndex        =   158
            Top             =   360
            Width           =   2295
            Begin VB.OptionButton RptSimple 
               Caption         =   "Artificial"
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
               Index           =   24
               Left            =   0
               TabIndex        =   40
               Tag             =   "19|3|-1"
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Mixta"
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
               Index           =   22
               Left            =   0
               TabIndex        =   39
               Tag             =   "19|2|-1"
               Top             =   360
               Width           =   1455
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "LME"
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
               Index           =   23
               Left            =   0
               TabIndex        =   38
               Tag             =   "19|1|-1"
               Top             =   0
               Width           =   1455
            End
         End
         Begin VB.Frame frSuplemento 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2160
            TabIndex        =   157
            Top             =   2280
            Width           =   2295
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   26
               Left            =   0
               TabIndex        =   42
               Tag             =   "21|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   25
               Left            =   720
               TabIndex        =   43
               Tag             =   "21|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox RptAlfa 
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
            Index           =   0
            Left            =   2160
            MaxLength       =   255
            TabIndex        =   41
            Tag             =   "20|1|-1"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label LblEtiqueta 
            Caption         =   "Suplemento de Fe < 2 años"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   29
            Left            =   120
            TabIndex        =   147
            Top             =   2160
            Width           =   1290
         End
         Begin VB.Label LblEtiqueta 
            Caption         =   "Inicio de Alimentación Complementaria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   28
            Left            =   120
            TabIndex        =   146
            Top             =   1560
            Width           =   1875
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Primeros 6 meses:"
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
            Index           =   27
            Left            =   120
            TabIndex        =   145
            Top             =   360
            Width           =   1485
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "3. Patológicos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   0
         Left            =   -69720
         TabIndex        =   133
         Top             =   480
         Width           =   5055
         Begin VB.Frame frAlergia 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   180
            Top             =   2880
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   42
               Left            =   0
               TabIndex        =   58
               Tag             =   "29|1|6"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   41
               Left            =   720
               TabIndex        =   59
               Tag             =   "29|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frCirugia 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   179
            Top             =   2520
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   40
               Left            =   0
               TabIndex        =   56
               Tag             =   "28|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   39
               Left            =   720
               TabIndex        =   57
               Tag             =   "28|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frTransfu 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   178
            Top             =   2160
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   38
               Left            =   0
               TabIndex        =   54
               Tag             =   "27|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   37
               Left            =   720
               TabIndex        =   55
               Tag             =   "27|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frHospitalizacion 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   177
            Top             =   1800
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   36
               Left            =   0
               TabIndex        =   52
               Tag             =   "26|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   35
               Left            =   720
               TabIndex        =   53
               Tag             =   "26|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frInfecciones 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   176
            Top             =   1440
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   34
               Left            =   0
               TabIndex        =   50
               Tag             =   "25|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   33
               Left            =   720
               TabIndex        =   51
               Tag             =   "25|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frEpilepsia 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   175
            Top             =   1080
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   32
               Left            =   0
               TabIndex        =   48
               Tag             =   "24|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   31
               Left            =   720
               TabIndex        =   49
               Tag             =   "24|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frSOBA 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   174
            Top             =   720
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   30
               Left            =   0
               TabIndex        =   46
               Tag             =   "23|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   29
               Left            =   720
               TabIndex        =   47
               Tag             =   "23|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frTbc 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2280
            TabIndex        =   173
            Top             =   360
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   28
               Left            =   0
               TabIndex        =   44
               Tag             =   "22|1|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   27
               Left            =   720
               TabIndex        =   45
               Tag             =   "22|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox RptEspecifique 
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
            Index           =   6
            Left            =   120
            TabIndex        =   60
            Top             =   3240
            Width           =   4695
         End
         Begin VB.OptionButton RptSimple 
            Caption         =   "Si"
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
            Index           =   43
            Left            =   2280
            TabIndex        =   61
            Tag             =   "30|1|7"
            Top             =   3600
            Width           =   615
         End
         Begin VB.TextBox RptEspecifique 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   3960
            Width           =   3615
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Especifique:"
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
            Index           =   33
            Left            =   120
            TabIndex        =   143
            Top             =   3960
            Width           =   975
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "TBC"
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
            Index           =   34
            Left            =   120
            TabIndex        =   142
            Top             =   360
            Width           =   330
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "SOBA / Asma"
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
            Index           =   35
            Left            =   120
            TabIndex        =   141
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Epilepsia"
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
            Index           =   36
            Left            =   120
            TabIndex        =   140
            Top             =   1080
            Width           =   675
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Infecciones"
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
            Index           =   37
            Left            =   120
            TabIndex        =   139
            Top             =   1440
            Width           =   930
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Hospitalizaciones"
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
            Index           =   38
            Left            =   120
            TabIndex        =   138
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Transfusiones sang."
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
            Index           =   39
            Left            =   120
            TabIndex        =   137
            Top             =   2160
            Width           =   1605
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Cirugia"
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
            Index           =   40
            Left            =   120
            TabIndex        =   136
            Top             =   2520
            Width           =   525
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Alergia a medicamentos"
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
            Index           =   41
            Left            =   120
            TabIndex        =   135
            Top             =   2880
            Width           =   1935
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Otros antec."
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
            Index           =   42
            Left            =   120
            TabIndex        =   134
            Top             =   3600
            Width           =   1035
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "III.Vivienda/Saneamiento Básico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   5
         Left            =   -69720
         TabIndex        =   113
         Top             =   480
         Width           =   5055
         Begin VB.Frame frAgua 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   160
            Top             =   360
            Width           =   2775
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   64
               Left            =   0
               TabIndex        =   103
               Tag             =   "41|1|8"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   65
               Left            =   720
               TabIndex        =   104
               Tag             =   "41|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame frAguaPotable 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   159
            Top             =   1200
            Width           =   2415
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   66
               Left            =   0
               TabIndex        =   106
               Tag             =   "42|1|9"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   67
               Left            =   720
               TabIndex        =   107
               Tag             =   "42|2|-1"
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.TextBox RptEspecifique 
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
            Index           =   9
            Left            =   1800
            TabIndex        =   108
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox RptEspecifique 
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
            Index           =   8
            Left            =   1800
            TabIndex        =   105
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Especificar:"
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
            Index           =   57
            Left            =   120
            TabIndex        =   117
            Top             =   720
            Width           =   900
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Especificar:"
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
            Index           =   56
            Left            =   120
            TabIndex        =   116
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Desague"
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
            Index           =   55
            Left            =   120
            TabIndex        =   115
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Agua potable"
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
            Index           =   54
            Left            =   120
            TabIndex        =   114
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "II. Antecedentes Familiares"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   4
         Left            =   -74880
         TabIndex        =   82
         Top             =   480
         Width           =   5055
         Begin VB.Frame frHepatitis 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   190
            Top             =   3600
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   62
               Left            =   720
               TabIndex        =   101
               Tag             =   "40|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   63
               Left            =   0
               TabIndex        =   100
               Tag             =   "40|1|7|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   7
               ItemData        =   "ucHCAntecedentes.ctx":0070
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":0083
               TabIndex        =   102
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frDrogradiccion 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   189
            Top             =   3240
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   60
               Left            =   720
               TabIndex        =   98
               Tag             =   "39|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   61
               Left            =   0
               TabIndex        =   97
               Tag             =   "39|1|8|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   8
               ItemData        =   "ucHCAntecedentes.ctx":00AB
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":00BE
               TabIndex        =   99
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frAlcoholismo 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   188
            Top             =   2880
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   58
               Left            =   720
               TabIndex        =   95
               Tag             =   "38|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   59
               Left            =   0
               TabIndex        =   94
               Tag             =   "38|1|9|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   9
               ItemData        =   "ucHCAntecedentes.ctx":00E6
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":00F9
               TabIndex        =   96
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frViolenciaF 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   187
            Top             =   2520
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   56
               Left            =   720
               TabIndex        =   92
               Tag             =   "37|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   57
               Left            =   0
               TabIndex        =   91
               Tag             =   "37|1|10|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   10
               ItemData        =   "ucHCAntecedentes.ctx":0121
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":0134
               TabIndex        =   93
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frAlergiaAnt 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   186
            Top             =   2160
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   54
               Left            =   720
               TabIndex        =   79
               Tag             =   "36|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   55
               Left            =   0
               TabIndex        =   78
               Tag             =   "36|1|6|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   6
               ItemData        =   "ucHCAntecedentes.ctx":015C
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":016F
               TabIndex        =   90
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frEpilepsiaAnt 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   185
            Top             =   1800
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   52
               Left            =   720
               TabIndex        =   76
               Tag             =   "35|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   53
               Left            =   0
               TabIndex        =   75
               Tag             =   "35|1|5|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   5
               ItemData        =   "ucHCAntecedentes.ctx":0197
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":01AA
               TabIndex        =   77
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frDiabetes 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   184
            Top             =   1440
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   50
               Left            =   720
               TabIndex        =   73
               Tag             =   "34|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   51
               Left            =   0
               TabIndex        =   72
               Tag             =   "34|1|4|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   4
               ItemData        =   "ucHCAntecedentes.ctx":01D2
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":01E5
               TabIndex        =   74
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frVIH 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   183
            Top             =   1080
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   48
               Left            =   720
               TabIndex        =   70
               Tag             =   "33|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   49
               Left            =   0
               TabIndex        =   69
               Tag             =   "33|1|3|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   3
               ItemData        =   "ucHCAntecedentes.ctx":020D
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":0220
               TabIndex        =   71
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frAsma 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   182
            Top             =   720
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   46
               Left            =   720
               TabIndex        =   67
               Tag             =   "32|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   47
               Left            =   0
               TabIndex        =   66
               Tag             =   "32|1|2|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   2
               ItemData        =   "ucHCAntecedentes.ctx":0248
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":025B
               TabIndex        =   68
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Frame frTuberculosis 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1800
            TabIndex        =   181
            Top             =   360
            Width           =   3135
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   44
               Left            =   720
               TabIndex        =   64
               Tag             =   "31|2|-1"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   45
               Left            =   0
               TabIndex        =   63
               Tag             =   "31|1|1|2"
               Top             =   0
               Width           =   615
            End
            Begin VB.ComboBox RptSimpleCombo 
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
               Index           =   1
               ItemData        =   "ucHCAntecedentes.ctx":0283
               Left            =   1440
               List            =   "ucHCAntecedentes.ctx":0296
               TabIndex        =   65
               Top             =   0
               Width           =   1575
            End
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Quién"
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
            Index           =   53
            Left            =   3240
            TabIndex        =   112
            Top             =   120
            Width           =   480
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Hepat.B"
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
            Index           =   52
            Left            =   120
            TabIndex        =   111
            Top             =   3600
            Width           =   660
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Drogadicción"
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
            Index           =   51
            Left            =   120
            TabIndex        =   110
            Top             =   3240
            Width           =   1035
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Alcoholismo"
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
            Index           =   50
            Left            =   120
            TabIndex        =   109
            Top             =   2880
            Width           =   945
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Violencia familiar"
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
            Index           =   49
            Left            =   120
            TabIndex        =   89
            Top             =   2520
            Width           =   1305
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Alergia a medicinas"
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
            Index           =   48
            Left            =   120
            TabIndex        =   88
            Top             =   2160
            Width           =   1530
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Epilepsia"
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
            Index           =   47
            Left            =   120
            TabIndex        =   87
            Top             =   1800
            Width           =   675
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Diabetes"
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
            Index           =   46
            Left            =   120
            TabIndex        =   86
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "VIH-SIDA."
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
            Index           =   45
            Left            =   120
            TabIndex        =   85
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "ASMA"
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
            Index           =   44
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   480
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Tuberculosis"
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
            Index           =   43
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "1.2 Parto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   5280
         TabIndex        =   81
         Top             =   480
         Width           =   5055
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   1455
            Index           =   7
            Left            =   120
            TabIndex        =   152
            Top             =   2280
            Width           =   4695
            Begin VB.OptionButton RptSimple 
               Caption         =   "Profesional de Salud"
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
               Index           =   9
               Left            =   1320
               TabIndex        =   15
               Tag             =   "7|1|-1"
               Top             =   0
               Width           =   1935
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Técnico"
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
               Index           =   10
               Left            =   3240
               TabIndex        =   16
               Tag             =   "7|2|-1"
               Top             =   0
               Width           =   1215
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "ACS"
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
               Index           =   11
               Left            =   1320
               TabIndex        =   17
               Tag             =   "7|3|-1"
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Familiar"
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
               Index           =   12
               Left            =   3240
               TabIndex        =   18
               Tag             =   "7|4|-1"
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Otro"
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
               Index           =   13
               Left            =   1320
               TabIndex        =   19
               Tag             =   "7|5|3"
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox RptEspecifique 
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
               Index           =   3
               Left            =   1320
               TabIndex        =   20
               Top             =   960
               Width           =   3255
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Atendido por:"
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
               Index           =   12
               Left            =   0
               TabIndex        =   172
               Top             =   0
               Width           =   1140
            End
         End
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   151
            Top             =   1800
            Width           =   4815
            Begin VB.OptionButton RptSimple 
               Caption         =   "EESS"
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
               Index           =   4
               Left            =   1320
               TabIndex        =   12
               Tag             =   "6|1|-1"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Domicilio"
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
               Index           =   5
               Left            =   2160
               TabIndex        =   13
               Tag             =   "6|2|-1"
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Consult.Partic."
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
               Index           =   6
               Left            =   3240
               TabIndex        =   14
               Tag             =   "6|3|-1"
               Top             =   0
               Width           =   1575
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Lugar del parto"
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
               Index           =   11
               Left            =   0
               TabIndex        =   171
               Top             =   0
               Width           =   1245
            End
         End
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   1215
            Index           =   5
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   4815
            Begin VB.TextBox RptEspecifique 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   11
               Top             =   600
               Width           =   4695
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Parto Eutócico"
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
               Index           =   7
               Left            =   1320
               TabIndex        =   9
               Tag             =   "5|1|-1"
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Complicado"
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
               Index           =   8
               Left            =   3120
               TabIndex        =   10
               Tag             =   "5|2|2"
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Parto"
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
               Index           =   9
               Left            =   0
               TabIndex        =   170
               Top             =   0
               Width           =   435
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Complicaciones del parto"
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
               Index           =   10
               Left            =   0
               TabIndex        =   169
               Top             =   360
               Width           =   2010
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "1.1 Embarazo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   5055
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   4
            Left            =   120
            TabIndex        =   166
            Top             =   2640
            Width           =   4815
            Begin VB.TextBox RptEspecifique 
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
               Index           =   1
               Left            =   1680
               TabIndex        =   8
               Top             =   360
               Width           =   3015
            End
            Begin VB.TextBox RptEntero 
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
               Index           =   2
               Left            =   1680
               TabIndex        =   7
               Tag             =   "4|1|1"
               Top             =   0
               Width           =   735
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Nº APN"
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
               Index           =   6
               Left            =   0
               TabIndex        =   168
               Top             =   0
               Width           =   615
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Lugar de APN"
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
               Index           =   7
               Left            =   0
               TabIndex        =   167
               Top             =   360
               Width           =   1125
            End
         End
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   163
            Top             =   1680
            Width           =   4815
            Begin VB.TextBox RptEntero 
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
               Index           =   1
               Left            =   1680
               TabIndex        =   4
               Tag             =   "2|1|-1"
               Top             =   0
               Width           =   735
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Nº de embarazo"
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
               Index           =   4
               Left            =   0
               TabIndex        =   164
               Top             =   0
               Width           =   1320
            End
         End
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   149
            Top             =   2160
            Width           =   4695
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   2
               Left            =   1680
               TabIndex        =   5
               Tag             =   "3|1|-1"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   3
               Left            =   2520
               TabIndex        =   6
               Tag             =   "3|2|-1"
               Top             =   0
               Width           =   735
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Atención Prenatal:"
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
               Index           =   5
               Left            =   0
               TabIndex        =   165
               Top             =   0
               Width           =   1515
            End
         End
         Begin VB.Frame frPregunta 
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   1
            Left            =   120
            TabIndex        =   148
            Top             =   360
            Width           =   4815
            Begin VB.TextBox RptEspecifique 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   0
               MultiLine       =   -1  'True
               TabIndex        =   3
               Top             =   600
               Width           =   4695
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Normal"
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
               Index           =   1
               Left            =   1680
               TabIndex        =   1
               Tag             =   "1|1|-1"
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "Complicado"
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
               Left            =   2880
               TabIndex        =   2
               Tag             =   "1|2|0"
               Top             =   0
               Width           =   1455
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Embarazo"
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
               Index           =   2
               Left            =   0
               TabIndex        =   162
               Top             =   0
               Width           =   780
            End
            Begin VB.Label LblEtiqueta 
               AutoSize        =   -1  'True
               Caption         =   "Patología(s) durante la gestación:"
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
               Index           =   3
               Left            =   0
               TabIndex        =   161
               Top             =   360
               Width           =   2745
            End
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Respiración y llanto al nacer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   3
         Left            =   -69720
         TabIndex        =   118
         Top             =   480
         Width           =   5055
         Begin VB.Frame frHospitalizacion 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   156
            Top             =   2640
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   20
               Left            =   0
               TabIndex        =   35
               Tag             =   "18|1|5"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   21
               Left            =   720
               TabIndex        =   36
               Tag             =   "18|2|-1"
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame frPatologiaNeonatal 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2280
            TabIndex        =   155
            Top             =   1800
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   18
               Left            =   0
               TabIndex        =   32
               Tag             =   "17|1|4"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   19
               Left            =   720
               TabIndex        =   33
               Tag             =   "17|2|-1"
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame frReannimacion 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2280
            TabIndex        =   154
            Top             =   1440
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   16
               Left            =   0
               TabIndex        =   30
               Tag             =   "16|1|-1"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   17
               Left            =   720
               TabIndex        =   31
               Tag             =   "16|2|-1"
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Frame frLlantoInmediato 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2280
            TabIndex        =   153
            Top             =   360
            Width           =   2535
            Begin VB.OptionButton RptSimple 
               Caption         =   "Si"
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
               Index           =   14
               Left            =   0
               TabIndex        =   26
               Tag             =   "13|1|-1"
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton RptSimple 
               Caption         =   "No"
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
               Index           =   15
               Left            =   720
               TabIndex        =   27
               Tag             =   "13|2|-1"
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.TextBox RptEspecifique 
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
            Index           =   5
            Left            =   2280
            TabIndex        =   37
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox RptEspecifique 
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
            Index           =   4
            Left            =   2280
            TabIndex        =   34
            Top             =   2160
            Width           =   2535
         End
         Begin VB.TextBox RptEntero 
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
            Index           =   4
            Left            =   2280
            TabIndex        =   29
            Tag             =   "15|1|-1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox RptEntero 
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
            Index           =   3
            Left            =   2280
            TabIndex        =   28
            Tag             =   "14|1|-1"
            Top             =   720
            Width           =   735
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Tiempo de hospitalización"
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
            Index           =   31
            Left            =   120
            TabIndex        =   126
            Top             =   3120
            Width           =   2085
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Hospitalización"
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
            Index           =   30
            Left            =   120
            TabIndex        =   125
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "APGAR 5 min"
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
            Index           =   26
            Left            =   120
            TabIndex        =   124
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Especifique:"
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
            Index           =   25
            Left            =   120
            TabIndex        =   123
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Patología Neonatal"
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
            Index           =   24
            Left            =   120
            TabIndex        =   122
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Reanimación"
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
            Index           =   23
            Left            =   120
            TabIndex        =   121
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "APGAR 1 min"
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
            Index           =   22
            Left            =   120
            TabIndex        =   120
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Inmediato"
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
            Index           =   21
            Left            =   120
            TabIndex        =   119
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "1.3 Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   2
         Left            =   -74880
         TabIndex        =   127
         Top             =   480
         Width           =   5055
         Begin VB.TextBox RptDouble 
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
            Index           =   5
            Left            =   2520
            TabIndex        =   21
            Tag             =   "8|1|-1"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox RptDouble 
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
            Index           =   4
            Left            =   2520
            TabIndex        =   25
            Tag             =   "12|1|-1"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox RptDouble 
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
            Index           =   3
            Left            =   2520
            TabIndex        =   24
            Tag             =   "11|1|-1"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox RptDouble 
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
            Index           =   2
            Left            =   2520
            TabIndex        =   23
            Tag             =   "10|1|-1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox RptDouble 
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
            Index           =   1
            Left            =   2520
            TabIndex        =   22
            Tag             =   "9|1|-1"
            Top             =   720
            Width           =   735
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Perímetro Torácico"
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
            Index           =   18
            Left            =   120
            TabIndex        =   132
            Top             =   1800
            Width           =   1545
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Perímetro cefálico"
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
            Index           =   17
            Left            =   120
            TabIndex        =   131
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Talla al nacer (cm)"
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
            Index           =   16
            Left            =   120
            TabIndex        =   130
            Top             =   1080
            Width           =   1500
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Peso al nacer (gr):"
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
            Index           =   15
            Left            =   120
            TabIndex        =   129
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label LblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Edad Gest. al nacer (sem):"
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
            Index           =   14
            Left            =   120
            TabIndex        =   128
            Top             =   360
            Width           =   2190
         End
      End
   End
End
Attribute VB_Name = "ucHCAntecedentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para registrar Antecedentes del Paciente
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Cadena As New sighentidades.cadena
Dim mo_Formulario As New sighentidades.Formulario
Dim lcBuscaParametro As New SIGHDatos.Parametros
'Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
'variables standares para el mantenimiento y auditoria
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long

Dim mb_ExistenDatos As Boolean

Dim mo_CmbIdTipoSexo As New sighentidades.ListaDespleglable
Dim ml_idPaciente As Long

'EVENTOS DEL CONTROL
Public Event SePresionoTeclaEspecial(KeyCode As Integer)

'=============================================================
'METODOS DE LECTURA
'=============================================================

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

'=============================================================
'METODOS DE LECTURA Y ESCRITURA
'=============================================================
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property

Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property

Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property


'Property Let idPaciente(lValue As Long)
'   ml_idPaciente = lValue
'End Property

Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property


'===================================================
'METODOS DE SOLO DE ESCRITURA
'===================================================


Property Let FechaNacimiento(lValue As Date)
'   ml_FechaNacimiento = lValue
End Property

Public Function inicializar(idPaciente As Long)
    ml_idPaciente = idPaciente
    sTabAntecedentes.Tab = 0
    BloqueoControlesIniciales
    Call cargarRespuestasPaciente
End Function

Private Sub CboRespuesta_Change(Index As Integer)

End Sub

Private Sub OptRespuesta_Click(Index As Integer)

End Sub

Private Sub RptAlfa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, RptAlfa
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub RptAlfa_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub RptDouble_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, RptEntero
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub RptDouble_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub RptEntero_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, RptEntero
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub RptEntero_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Sub RptEspecifique_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, RptEspecifique
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub RptEspecifique_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim VarAnularTecla As Boolean
    VarAnularTecla = False
    Select Case Index
        Case -1:
            'solo digitos
            If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
                VarAnularTecla = True
            End If
        Case -1:
            'numeros con decimal
            If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
                VarAnularTecla = True
            End If
        Case -1:
            'solo letras
            If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
                VarAnularTecla = True
            End If
        Case Else
            'alfanumerico
    End Select
    If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) And VarAnularTecla = True Then
        KeyAscii = 0
   End If
End Sub

Private Sub RptSimple_Click(Index As Integer)
    Select Case Index
        Case 0:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(0), RptSimple(Index).Value
        Case 1:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(0), Not RptSimple(Index).Value
        Case 2:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(1), RptSimple(Index).Value
            mo_Formulario.HabilitarDeshabilitar RptEntero(2), RptSimple(Index).Value
        Case 3:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(1), Not RptSimple(Index).Value
            mo_Formulario.HabilitarDeshabilitar RptEntero(2), Not RptSimple(Index).Value
        Case 8:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(2), RptSimple(Index).Value
        Case 7:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(2), Not RptSimple(Index).Value
        Case 13:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(3), RptSimple(Index).Value
        Case 9, 10, 11, 12:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(3), Not RptSimple(Index).Value
        Case 18:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(4), RptSimple(Index).Value
        Case 19:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(4), Not RptSimple(Index).Value
        Case 20:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(5), RptSimple(Index).Value
        Case 21:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(5), Not RptSimple(Index).Value
        Case 42:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(6), RptSimple(Index).Value
        Case 41:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(6), Not RptSimple(Index).Value
        Case 43:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(7), RptSimple(Index).Value
        Case 64:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(8), RptSimple(Index).Value
        Case 65:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(8), Not RptSimple(Index).Value
        Case 66:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(9), RptSimple(Index).Value
        Case 67:
            mo_Formulario.HabilitarDeshabilitar RptEspecifique(9), Not RptSimple(Index).Value
        Case 45:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(1), RptSimple(Index).Value
        Case 44:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(1), Not RptSimple(Index).Value
        Case 47:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(2), RptSimple(Index).Value
        Case 46:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(2), Not RptSimple(Index).Value
        Case 49:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(3), RptSimple(Index).Value
        Case 48:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(3), Not RptSimple(Index).Value
        Case 51:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(4), RptSimple(Index).Value
        Case 50:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(4), Not RptSimple(Index).Value
        Case 53:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(5), RptSimple(Index).Value
        Case 52:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(5), Not RptSimple(Index).Value
        Case 55:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(6), RptSimple(Index).Value
        Case 54:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(6), Not RptSimple(Index).Value
        Case 57:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(10), RptSimple(Index).Value
        Case 56:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(10), Not RptSimple(Index).Value
        Case 59:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(9), RptSimple(Index).Value
        Case 58:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(9), Not RptSimple(Index).Value
        Case 61:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(8), RptSimple(Index).Value
        Case 60:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(8), Not RptSimple(Index).Value
        Case 63:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(7), RptSimple(Index).Value
        Case 62:
            mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(7), Not RptSimple(Index).Value
    End Select
End Sub

Private Sub RptSimple_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, RptSimple
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub RptSimpleCombo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, RptSimpleCombo
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub RptSimpleCombo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub


'======================================================================
'METODOS
'======================================================================

Public Function SeRegistroAntecedentes() As Boolean
    Dim oReglasAntecedentePaciente As New ReglasAntecedentesPaciente
    Dim oGrupoHCPaciente As DOAtenInteGrupoHCPaciente
    Dim rs As ADODB.Recordset
    
    oGrupoHCPaciente.idPaciente = ml_idPaciente
    oGrupoHCPaciente.IdAtenInteGrupo = getIdAtenInteGrupo()
    SeRegistroAntecedentes = False
    
    Set rs = oReglasAntecedentePaciente.ListarPreguntasPorPacienteYGrupo(oGrupoHCPaciente)
    If Not (rs.EOF = True And rs.BOF = True) Then
        SeRegistroAntecedentes = True
    End If
End Function

Private Sub cargarRespuestasPaciente()
    Dim oReglasAntecedentePaciente As New ReglasAntecedentesPaciente
    Dim oGrupoHCPaciente As DOAtenInteGrupoHCPaciente
    Dim rsRespuestas As ADODB.Recordset
    Dim rsPreguntas As ADODB.Recordset
    Dim sTipoControl As String
    
    Dim arraySplitTagControl() As String
    Set oGrupoHCPaciente = New DOAtenInteGrupoHCPaciente
    oGrupoHCPaciente.idPaciente = ml_idPaciente
    oGrupoHCPaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    
    Set rsPreguntas = oReglasAntecedentePaciente.ListarPreguntasPorPacienteYGrupo(oGrupoHCPaciente)
    
    Set rsRespuestas = oReglasAntecedentePaciente.ListarRespuestasPorPacienteYGrupo(oGrupoHCPaciente)
    
    If oReglasAntecedentePaciente.MensajeError <> "" Then
        MsgBox oReglasAntecedentePaciente.MensajeError
        Exit Sub
    End If
    
    If Not (rsPreguntas.EOF = True And rsPreguntas.BOF = True) _
            And Not (rsRespuestas.EOF = True And rsRespuestas.BOF = True) Then
    
        Dim oControl As Control
        
        For Each oControl In UserControl.Controls
            If Left(UCase(oControl.Name), 3) = "RPT" And oControl.Tag <> "" Then
            
                arraySplitTagControl = Split(oControl.Tag, "|")
                
                rsPreguntas.MoveFirst
                rsPreguntas.Find "IdPregunta = " & Val(arraySplitTagControl(0))
                
                If rsPreguntas.EOF = False Then
                    rsRespuestas.Filter = "IdGrupoHCPaciente=" & rsPreguntas!IdGrupoHCPaciente & " AND ItemRespuesta=" & Val(arraySplitTagControl(1))
                    If Not (rsRespuestas.EOF = True And rsRespuestas.BOF = True) Then
                        
                        sTipoControl = UCase(TypeName(oControl))
                        
                        Select Case sTipoControl
                            Case "TEXTBOX":
                                oControl.Text = devuelveRespuesta(rsRespuestas)
                            Case "CHECKBOX":
                                oControl.Value = 1
                            Case "OPTIONBUTTON":
                                oControl.Value = True
                        End Select
                        
                        Call cargarValorEspecificacionRespuesta(rsRespuestas, arraySplitTagControl)
                    End If
                End If
            End If
        Next
    End If
End Sub

Public Function devuelveRespuesta(rsRespuestas As ADODB.Recordset) As String
    Select Case rsRespuestas!TipoValorRespuesta
        Case sighTipoDatoRespuesta.Numerico:
            devuelveRespuesta = IIf(IsNull(rsRespuestas!ValorNumero), "", rsRespuestas!ValorNumero)
        Case sighTipoDatoRespuesta.fecha:
            devuelveRespuesta = IIf(IsNull(rsRespuestas!ValorFecha), "", rsRespuestas!ValorFecha)
        Case Else
            devuelveRespuesta = IIf(IsNull(rsRespuestas!ValorTexto), "", rsRespuestas!ValorTexto)
    End Select
End Function


Public Function asignaValorRespuesta(ByRef oDORespuestaPaciente As DOAtenInteHCRespuestaPaciente, _
        sRespuesta As String) As Boolean
            
    oDORespuestaPaciente.ValorFecha = 0
    oDORespuestaPaciente.ValorFechaFin = 0
    oDORespuestaPaciente.ValorNumero = 0
    oDORespuestaPaciente.ValorNumeroFin = 0
    oDORespuestaPaciente.ValorTexto = ""
    
    Dim oReglasAntecedentePaciente As New ReglasAntecedentesPaciente
    Dim oDOPregunta As New DOAtenIntePregunta
    
    Set oDOPregunta = oReglasAntecedentePaciente.PreguntaSeleccionarPorId(oDORespuestaPaciente.IdPregunta)
    
    
    Select Case oDOPregunta.TipoValorRespuesta
        Case sighTipoDatoRespuesta.Numerico:
            oDORespuestaPaciente.ValorNumero = CDbl(sRespuesta)
        Case sighTipoDatoRespuesta.fecha:
            oDORespuestaPaciente.ValorFecha = CDate(sRespuesta)
        Case Else
            oDORespuestaPaciente.ValorTexto = sRespuesta
    End Select
End Function

Public Function asignaValorEspecificacionRespuesta(ByRef oDORespuestaPaciente As DOAtenInteHCRespuestaPaciente, _
                    arraySplitTag() As String) As Boolean
    If Val(arraySplitTag(2)) = -1 Then
        oDORespuestaPaciente.ValorEspecificacion = ""
    Else
        If UBound(arraySplitTag) = 2 Then
            oDORespuestaPaciente.ValorEspecificacion = RptEspecifique(Val(arraySplitTag(2))).Text
        Else
            'extraer el otros controles por implementar
            'Implemete un tipado de controles en la propiedad tag
            Select Case Val(arraySplitTag(3))
                Case 2:
                    oDORespuestaPaciente.ValorEspecificacion = RptSimpleCombo(Val(arraySplitTag(2))).Text
                    
                Case Else
                    MsgBox "Implemete un tipado de controles en la propiedad tag"
            End Select
        End If
    End If
End Function

Public Function cargarValorEspecificacionRespuesta(ByRef rsRespuesta As ADODB.Recordset, _
                    arraySplitTag() As String) As Boolean
    If Val(arraySplitTag(2)) <> -1 Then
        If UBound(arraySplitTag) = 2 Then
            RptEspecifique(Val(arraySplitTag(2))).Text = IIf(IsNull(rsRespuesta!ValorEspecificacion), "", rsRespuesta!ValorEspecificacion)
        Else
            'extraer el otros controles por implementar
            'Implemete un tipado de controles en la propiedad tag
            Select Case Val(arraySplitTag(3))
                Case 2:
                    RptSimpleCombo(arraySplitTag(2)).Text = IIf(IsNull(rsRespuesta!ValorEspecificacion), "", rsRespuesta!ValorEspecificacion)
                Case Else
                    MsgBox "Implemete un tipado de controles en la propiedad tag"
            End Select
            
            
            
        End If
    End If
End Function

Public Function SeGeneroPlanIntegral() As Boolean
    Dim oReglasAntecedentesPaciente As New ReglasAntecedentesPaciente
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ml_idPaciente
    If oReglasAtencionIntegral.SeleccionarPlanIntegralPorPacienteYGrupo(oDOAtenIntePlanIntePaciente) Is Nothing Then
        ms_MensajeError = oReglasAtencionIntegral.MensajeError
        SeGeneroPlanIntegral = False
    Else
        SeGeneroPlanIntegral = True
    End If
End Function

'Public Function generarPlanAtencionIntegral()
'
'End Function


Public Function grabarAntecedentePaciente() As Boolean
    If ValidarDatosIngreso = False Then
        Exit Function
    End If
    
    Dim oArrayPreguntas() As Long
    
    Dim oReglasAntecedentePaciente As New ReglasAntecedentesPaciente
    Dim oGrupoHCPaciente As DOAtenInteGrupoHCPaciente
    Dim oDORespuestaPaciente As DOAtenInteHCRespuestaPaciente
    
    Dim cPreguntas As New Collection
    Dim cRespuestas As New Collection
    Dim agregarRespuesta As Boolean
    Dim Respuesta As String, sTipoControl As String
    Dim arraySplitTagControl() As String
    
    Dim oControl As Control
    ReDim Preserve oArrayPreguntas(0)
    
    For Each oControl In UserControl.Controls
        agregarRespuesta = False
        If Left(UCase(oControl.Name), 3) = "RPT" Then
            If oControl.Tag <> "" Then
                
            arraySplitTagControl = Split(oControl.Tag, "|")
            
            sTipoControl = UCase(TypeName(oControl))
                    
            Select Case sTipoControl
                Case "TEXTBOX":
                    If Trim(oControl.Text) <> "" Then
                        agregarRespuesta = True
                        Respuesta = Trim(oControl.Text)
                    End If
                    
                Case "CHECKBOX":
                    If oControl.Value = 1 Then
                        agregarRespuesta = True
                        Respuesta = Trim(oControl.Caption)
                    End If
                Case "OPTIONBUTTON":
                    If oControl.Value = True Then
                        agregarRespuesta = True
                        Respuesta = Trim(oControl.Caption)
                    End If
            End Select
                        
            If agregarRespuesta = True Then
                ReDim Preserve oArrayPreguntas(UBound(oArrayPreguntas) + 1)
                If buscarPreguntaEnArray(oArrayPreguntas, arraySplitTagControl(0)) = False Then
                    Set oGrupoHCPaciente = New DOAtenInteGrupoHCPaciente
                    oGrupoHCPaciente.IdAtenInteGrupo = getIdAtenInteGrupo()
                    oGrupoHCPaciente.idPaciente = ml_idPaciente
                    oGrupoHCPaciente.IdPregunta = arraySplitTagControl(0)
                    oGrupoHCPaciente.IdUsuarioAuditoria = ml_IdUsuario
                    cPreguntas.Add oGrupoHCPaciente
                    oArrayPreguntas(UBound(oArrayPreguntas) - 1) = oGrupoHCPaciente.IdPregunta
                End If
                Set oDORespuestaPaciente = New DOAtenInteHCRespuestaPaciente
                oDORespuestaPaciente.EsActivo = True
                oDORespuestaPaciente.idPaciente = ml_idPaciente
                oDORespuestaPaciente.IdPregunta = arraySplitTagControl(0)
                oDORespuestaPaciente.IdUsuarioAuditoria = ml_IdUsuario
                oDORespuestaPaciente.ItemRespuesta = arraySplitTagControl(1)
                Call asignaValorEspecificacionRespuesta(oDORespuestaPaciente, arraySplitTagControl)
                Call asignaValorRespuesta(oDORespuestaPaciente, Respuesta)
                cRespuestas.Add oDORespuestaPaciente
                
            End If
            End If
        End If
    Next
    grabarAntecedentePaciente = oReglasAntecedentePaciente.grabarRespuestasPaciente(cPreguntas, cRespuestas)
End Function


Public Function ValidarDatosIngreso() As Boolean

    ValidarDatosIngreso = True
End Function


Public Function buscarPreguntaEnArray(oArrayPreguntas() As Long, valorBuscado As String) As Boolean
    buscarPreguntaEnArray = False
    Dim i As Integer
    For i = 0 To UBound(oArrayPreguntas) - 1
        If oArrayPreguntas(i) = Val(valorBuscado) Then
            buscarPreguntaEnArray = True
            Exit For
        End If
    Next
End Function

Private Function getIdAtenInteGrupo() As Long
    getIdAtenInteGrupo = sighGrupoEdad.Nino
End Function

Private Sub BloqueoControlesIniciales()
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(1), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(0), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(2), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(3), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(4), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(5), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(6), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(7), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(8), False
    mo_Formulario.HabilitarDeshabilitar RptEspecifique(9), False
    mo_Formulario.HabilitarDeshabilitar RptEntero(2), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(1), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(2), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(3), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(4), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(5), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(6), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(7), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(8), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(9), False
    mo_Formulario.HabilitarDeshabilitar RptSimpleCombo(10), False
    
End Sub
