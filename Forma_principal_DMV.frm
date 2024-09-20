VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Forma_principal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "             Vehicle Moving Permit"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14490
   ControlBox      =   0   'False
   Icon            =   "Forma_principal_DMV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "1"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   6000
      Width           =   470
   End
   Begin Project1.lvButtons_H btnsearch 
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":377EE
      ImgSize         =   48
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   9
      Top             =   2760
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":3831E
      ImgSize         =   40
      cBack           =   14737632
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":38C80
      ImgSize         =   40
      cBack           =   14737632
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Index           =   4
      Left            =   5400
      TabIndex        =   23
      Top             =   3600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":395E2
      ImgSize         =   40
      cBack           =   14737632
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   20
      Top             =   2760
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":39F44
      ImgSize         =   40
      cBack           =   14737632
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   21
      Top             =   3600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":3A8A6
      ImgSize         =   40
      cBack           =   14737632
   End
   Begin VB.TextBox txtvin 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   740
      TabIndex        =   4
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "NONE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtyear 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtmodel 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtcust_id 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   11400
      ScaleHeight     =   6795
      ScaleWidth      =   7275
      TabIndex        =   15
      Top             =   960
      Width           =   7335
   End
   Begin Project1.lvButtons_H btnDialogo 
      Height          =   375
      Left            =   11400
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Carga"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboimpre 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   360
      TabIndex        =   11
      Top             =   5640
      Width           =   3855
   End
   Begin VB.TextBox txtmake 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtlicense 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   740
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin Project1.lvButtons_H btnimprime 
      Height          =   615
      Left            =   4320
      TabIndex        =   12
      Top             =   5640
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":3B208
      ImgSize         =   48
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   4680
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   661
      Caption         =   "&Exit"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   65280
      cGradient       =   65280
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   8421504
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   25
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":3CDF1
      ImgSize         =   40
      cBack           =   14737632
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   3000
      TabIndex        =   27
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
      BackColorFixed  =   8421504
      ForeColorFixed  =   14737632
      ScrollBars      =   2
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin Project1.lvButtons_H btnlimpia 
      Height          =   615
      Left            =   9000
      TabIndex        =   33
      Top             =   2760
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_principal_DMV.frx":3D753
      ImgSize         =   48
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "2"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "3"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   38
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "4"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   4
      Left            =   7080
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "5"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   40
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "6"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   6
      Left            =   7560
      TabIndex        =   41
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "7"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   7
      Left            =   7800
      TabIndex        =   42
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "8"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   43
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "9"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   9
      Left            =   8280
      TabIndex        =   44
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "10"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   10
      Left            =   8520
      TabIndex        =   45
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "11"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   11
      Left            =   8760
      TabIndex        =   46
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "12"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   12
      Left            =   9000
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "13"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   13
      Left            =   9240
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
      Caption         =   "14"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnfila 
      Height          =   255
      Index           =   14
      Left            =   9480
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Caption         =   "15"
      CapAlign        =   2
      BackStyle       =   7
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   5880
      Picture         =   "Forma_principal_DMV.frx":3E4DB
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   8760
      Top             =   2640
      Width           =   15
   End
   Begin VB.Label Label5 
      Caption         =   "Version 1.14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3960
      TabIndex        =   49
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Copies:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (C) 2024"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   30
      Top             =   6550
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Created by: Hector Navarro and Cintia Cadena"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   29
      Top             =   6360
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the row to select the VIN number:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   6
      Left            =   3000
      TabIndex        =   28
      Top             =   240
      Width           =   2985
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " (If any):"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   720
      TabIndex        =   26
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   5760
      TabIndex        =   22
      Top             =   3360
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Identification Number (VIN) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   720
      TabIndex        =   19
      Top             =   3360
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   5760
      TabIndex        =   18
      Top             =   2520
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   960
      Picture         =   "Forma_principal_DMV.frx":46E9C
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   16
      Top             =   320
      Width           =   1440
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make of Vehicle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle license number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   1980
   End
End
Attribute VB_Name = "Forma_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer



'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


' PARA REDONDEAR LA FORMA
' Crea la región
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal X3 As Long, _
    ByVal Y3 As Long) As Long
  
'Establece la región
Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
  

'  ---- esto es para el texto vertical

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private fnt As CLogFont

Const LF_FACESIZE = 32


Public Function GetIPHostName() As String
On Error Resume Next
    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function

Public Function HiByte(ByVal wParam As Integer) As Byte
  On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function
Public Function LoByte(ByVal wParam As Integer) As Byte
On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function

Public Sub SocketsCleanup()
On Error Resume Next
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Public Function SocketsInitialize() As Boolean
On Error Resume Next

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Public Sub enca()
'On Error Resume Next


If Grid2.Rows = 1 Then Grid2.Rows = 2

Grid2.ColWidth(0) = 400
Grid2.ColAlignment(0) = flexAlignLeftCenter


Grid2.ColWidth(1) = 800 ' ID
Grid2.ColAlignment(1) = flexAlignRightCenter

Grid2.ColWidth(2) = 1200   ' License plate
Grid2.ColAlignment(2) = flexAlignCenterCenter

Grid2.ColWidth(3) = 2200   ' VIN
Grid2.ColAlignment(3) = flexAlignCenterCenter

Grid2.ColWidth(4) = 1500   ' Make
Grid2.ColAlignment(4) = flexAlignCenterCenter

Grid2.ColWidth(5) = 740   ' Year
Grid2.ColAlignment(5) = flexAlignCenterCenter

Grid2.ColWidth(6) = 2200   ' Model
Grid2.ColAlignment(6) = flexAlignCenterCenter




Grid2.Row = 0

Grid2.Col = 1
Grid2.ColAlignment(1) = flexAlignCenterCenter
Grid2.Text = "ID"

Grid2.Col = 2
Grid2.ColAlignment(2) = flexAlignCenterCenter
Grid2.Text = "License Plate"


Grid2.Col = 3
Grid2.ColAlignment(3) = flexAlignCenterCenter
Grid2.Text = "VIN #"


Grid2.Col = 4
Grid2.ColAlignment(4) = flexAlignCenterCenter
Grid2.Text = "Make"


Grid2.Col = 5
Grid2.ColAlignment(5) = flexAlignCenterCenter
Grid2.Text = "Year"


Grid2.Col = 6
Grid2.ColAlignment(6) = flexAlignCenterCenter
Grid2.Text = "Model"



For t = 1 To Grid2.Rows - 1
   Grid2.Row = t
   Grid2.Col = 0
   Grid2.Text = t
Next t


Grid2.FixedRows = 1
Grid2.FixedCols = 1

Grid2.Row = 1
Grid2.Col = 1
End Sub
' ----------
Private Sub Redondear_Formulario(El_Form As Form, Radio As Long)
  
Dim Region As Long
Dim ret As Long
Dim Ancho As Long
Dim Alto As Long
Dim old_Scale As Integer
      
    ' guardar la escala
    old_Scale = El_Form.ScaleMode
      
    ' cambiar la escala a pixeles
    El_Form.ScaleMode = vbPixels
      
    'Obtenemos el ancho y alto de la region del Form
    Ancho = El_Form.ScaleWidth
    Alto = El_Form.ScaleHeight
  
    'Pasar el ancho alto del formualrio y el valor de redondeo .. es decir el radio
    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
  
    ' Aplica la región al formulario
    ret = SetWindowRgn(El_Form.hWnd, Region, True)
      
    ' restaurar la escala
    El_Form.ScaleMode = old_Scale
  
End Sub
Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
 
 
 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE"
 user_ini$ = "payroll"
 bd_ini$ = "laesystemja"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com"   ' "167.114.199.93"  '

 

 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub
Public Sub carga_impresoras()
On Error Resume Next

Dim cImprGen As String
    cImprGen = cboimpre.Text
    
cboimpre.Clear
ruta$ = "c:\dmv\"
    
If Dir$(ruta$ + "printer") <> "" Then
 nf = FreeFile
 Open ruta$ + "printer" For Input Shared As #nf
 Lock #nf
 Line Input #nf, P1$
 Line Input #nf, P2$
 Unlock #nf
 Close #nf
 
 cImprGen = P1$
 cboimpre.Text = P1$

End If
    
    
    
    
For Each xprint In Printers
           If xprint.DeviceName = cImprGen Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next
        
        
        
For Each xprint In Printers
        cboimpre.AddItem xprint.DeviceName
Next
        
        
nf = FreeFile
 Open ruta$ + "printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 
 
 For t = 0 To cboimpre.ListCount - 1
   If cboimpre.List(t) = Printer.DeviceName Then
       cboimpre.ListIndex = t
       Exit For
   End If
 Next t
        
        
        
        
End Sub

Private Sub btnborrar1_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
  txtlicense.Text = ""
  txtlicense.SetFocus
Case 1
  txtmake.Text = ""
  txtmake.SetFocus
Case 2
  txtmodel.Text = ""
  txtmodel.SetFocus
Case 3
  txtvin.Text = ""
  txtvin.SetFocus
Case 4
  txtyear.Text = ""
  txtyear.SetFocus
Case 5
  txtcust_id.Text = ""
  txtcust_id.SetFocus
  
End Select


End Sub

Private Sub btnDialogo_Click()
On Error Resume Next

Dim BeginPage, EndPage, NumCopies, Orientation, i
' Set Cancel to True.








CommonDialog1.CancelError = True
'On Error GoTo ErrHandler
' Display the Print dialog box.
Printer.PaperBin = 2
Printer.Orientation = 2

CommonDialog1.Flags = &H40 + &H100000 + &H200000 + &H80000 '+ &H400
CommonDialog1.ShowPrinter




' Get user-selected values from the dialog box.
BeginPage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
Orientation = CommonDialog1.Orientation
Printer.EndDoc


For i = 1 To NumCopies
' Put code here to send data to yourprinter.
Next
End Sub

Private Sub btnfila_Click(Index As Integer)
On Error Resume Next


a = Index + 1
Grid2.Row = a
Grid2.Col = 2
txtlicense.Text = Grid2.Text

Grid2.Col = 3
txtvin.Text = Grid2.Text

Grid2.Col = 4
txtmake.Text = Grid2.Text

Grid2.Col = 5
txtyear.Text = Grid2.Text

Grid2.Col = 6
txtmodel.Text = Grid2.Text
txtlicense.SetFocus


If Grid2.Rows <= 2 Then
   btnfila(0).Value = False
End If

End Sub

Private Sub btnimprime_Click()
On Error Resume Next

r$ = MsgBox("Do you wish to print the form?", 4, "Attention")
If r$ = "7" Then Exit Sub


For Y = 1 To Combo1.List(Combo1.ListIndex)

Picture1_Click

' Set the PictureBox's ScaleMode to pixels to
    ' make things interesting.
    Picture1.ScaleMode = vbPixels

    ' Print the picture.
    Printer.PaintPicture Picture1.Image, 0, 0

    ' Get the picture's dimensions in the printer's scale
    ' mode.
    wid = ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, Printer.ScaleMode)
    hgt = ScaleY(Picture1.ScaleHeight, Picture1.ScaleMode, Printer.ScaleMode)

    ' Draw the box.
    'Printer.Line (1440, 1440)-Step(wid, hgt), , B

    ' Finish printing.
    Printer.EndDoc
    
Next Y
    
    

End Sub

Private Sub btnlimpia_Click()
On Error Resume Next
txtcust_id.Text = ""
txtlicense.Text = ""
txtmake.Text = ""
txtmodel.Text = ""
txtvin.Text = ""
txtyear.Text = ""
Grid2.Rows = 2
For t = 0 To Grid2.Rows - 1
   Grid2.Row = t
   For Y = 0 To Grid2.cols - 1
      Grid2.Col = Y
      Grid2.Text = ""
   Next Y
Next t

Combo1.ListIndex = 0

For t = 0 To 14
 btnfila(t).Visible = False
Next t
txtcust_id.SetFocus

End Sub

Private Sub btnok_Click()
On Error Resume Next

 base.Close
 End
 
End Sub

Private Sub btnprint_Click()




End Sub

Private Sub btnsearch_Click()
On Error Resume Next
 
 
If txtcust_id.Text = "" Then
   Exit Sub
End If
g$ = txtcust_id.Text

btnlimpia_Click

txtcust_id.Text = g$



Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
    
    
    sSelect = "select IDVehicleInsured , LicensePlate, VINNumber, Make, Year, model from VehicleInsured vehicle " & _
              "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=vehicle.IdPolicieHDR " & _
              "Where IdCustomer = '" + txtcust_id.Text + "'"
    

   ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    For t = 0 To 14
      btnfila(t).Visible = False
    Next t
    
    For t = 1 To Grid2.Rows - 1
      btnfila(t - 1).Visible = True
      btnfila(t - 1).Value = False
    Next t
    
    
    enca
    
    

End Sub

Private Sub cboimpre_Click()
On Error Resume Next


For Each xprint In Printers
           If xprint.DeviceName = cboimpre.Text Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next


nf = FreeFile
 Open "c:\dmv\printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
End Sub


Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
   txtvin.Text = "NONE"
   txtvin.Enabled = False
  
Else
   txtvin.Text = ""
   txtvin.Enabled = True
   
End If

 txtvin.SetFocus
 
 
End Sub

Private Sub Form_Load()
On Error Resume Next

If (App.PrevInstance = True) Then
  'base.Close
  End
End If


 a$ = GetIPHostName()

  nf = FreeFile
  Open "\\192.168.84.215\dmv\" + a$ + "-in" For Output Shared As #nf
  Lock #nf
  Print #nf, Format(Now, "mm/dd/yyyy  hh:mm am/pm")
  Unlock #nf
  Close #nf
  



' verifica si hay actualizacion
nf = FreeFile
  Open "\\192.168.84.215\dmv\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_actual$
  Unlock #nf
  Close #nf
  
  nf = FreeFile
  Open "c:\dmv\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_programa$
  Unlock #nf
  Close #nf
  
  If Val(version_programa$) < Val(version_actual$) Then
     actualiza = 1
     r$ = Shell("\\192.168.84.215\dmv\actualizador.exe", vbNormalFocus)
     
     Hide
     Refresh
     End
     
  End If
  
  


MkDir "c:\dmv"
carga_impresoras
Conecta_SQL

Set fnt = New CLogFont
Set fnt.LOGFONT = Picture1.Font
fnt.Rotation = 90

 Top = 200
 Left = (Screen.Width - Width) / 2
   
   
 Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
      'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1366 '1024
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 1024 '940 '1024
      'End If
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = (Xpixels / DesignX)  ' 0.78
        ScaleFactorY = (Ypixels / DesignY)  ' 0.78
      Else
        ScaleFactorX = (Xpixels / DesignX)
        ScaleFactorY = (Ypixels * 1.3 / DesignY)
      
        'ScaleFactorX = 1360 / DesignX
        'ScaleFactorY = 1024 / DesignY
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        forma_main.Height = 9000 'Me.Height ' Remember the current size
        forma_main.Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0


 ' Le pasamos el formulario y el radio de redondeo de la forma
    Call Redondear_Formulario(Me, 100)
   
   
Combo1.Clear
For t = 1 To 30
   Combo1.AddItem t
Next t
   
Combo1.ListIndex = 0

End Sub


Private Sub Form_Resize()
 On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1
End Sub




Public Sub Impresion_texto()

End Sub

Private Sub Grid2_EnterCell()
On Error Resume Next
a = Grid2.Row
Grid2.Col = 2
txtlicense.Text = Grid2.Text

Grid2.Col = 3
txtvin.Text = Grid2.Text

Grid2.Col = 4
txtmake.Text = Grid2.Text

Grid2.Col = 5
txtyear.Text = Grid2.Text

Grid2.Col = 6
txtmodel.Text = Grid2.Text

txtlicense.SetFocus

End Sub




Private Sub Picture1_Click()
On Error Resume Next
Dim hfont As Long
Dim wid As Single
Dim hgt As Single

' linea License Number
With Picture1
  hfont = SelectObject(.hdc, fnt.Handle)
  .CurrentX = 2100  'X
  .CurrentY = 7150  'Y
  Picture1.Print txtlicense.Text  ' "1. License Number"

  Call SelectObject(.hdc, hfont)
End With


' linea 2 Make of vehicle
With Picture1
  hfont = SelectObject(.hdc, fnt.Handle)
  .CurrentX = 2100  'X
  .CurrentY = 3700 '3820
  Picture1.Print txtmake.Text  '"2. Make"

  Call SelectObject(.hdc, hfont)
End With



' linea 3 Model
With Picture1
  hfont = SelectObject(.hdc, fnt.Handle)
  .CurrentX = 2100  'X
  .CurrentY = 1800 '2000
  Picture1.Print txtmodel.Text  ' "3. Model"

  Call SelectObject(.hdc, hfont)
End With







' linea 4 VIN Number
With Picture1
  hfont = SelectObject(.hdc, fnt.Handle)
  .CurrentX = 2580  '2650  'X
  .CurrentY = 7150
  Picture1.Print txtvin.Text  '"4. VIN Number"

  Call SelectObject(.hdc, hfont)
End With




' linea 5 Year
With Picture1
  hfont = SelectObject(.hdc, fnt.Handle)
  .CurrentX = 2580 '2650  'X
  .CurrentY = 1800 '2000
  Picture1.Print txtyear.Text  '"5. Year"

  Call SelectObject(.hdc, hfont)
End With



End Sub

Private Sub txtcust_id_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   btnsearch_Click
End If

End Sub


Private Sub txtyear_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
   Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   KeyAscii = 0
End If

End Sub


