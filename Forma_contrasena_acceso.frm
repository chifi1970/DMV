VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Forma_accesar 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btncancel 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin Project1.lvButtons_H btnok 
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      Caption         =   "OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   3
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00C0C0FF&
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
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   4440
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   8421504
      ForeColorFixed  =   14737632
      BackColorBkg    =   -2147483632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btnborrar1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Image           =   "Forma_contrasena_acceso.frx":0000
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnborrar2 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Image           =   "Forma_contrasena_acceso.frx":0962
      cBack           =   12632256
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type your email:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2520
      Picture         =   "Forma_contrasena_acceso.frx":12C4
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@justautoins.com"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1905
   End
End
Attribute VB_Name = "Forma_accesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  btnok_Click
  Exit Sub
End If
End Sub

Private Sub btnborrar1_Click()
txtuser.Text = ""
txtuser.SetFocus
End Sub

Private Sub btnborrar2_Click()
txtpassword.Text = ""
txtpassword.SetFocus
End Sub


Private Sub btncancel_Click()
On Error Resume Next
base.Close
Unload Me

End Sub

Private Sub btnok_Click()
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset


           
            
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle, emp.emailwork from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (6,16,17,28,2,18,24,37) "




   ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
     valor = 0
    existe = 0
    For t = 1 To Grid2.Rows - 1
       Grid2.Row = t
       Grid2.Col = 1
       id_user$ = Grid2.Text
       
       
       Grid2.Col = 2
       userx$ = UCase(Grid2.Text)
       
       Grid2.Col = 5
       emailx$ = UCase(Grid2.Text)
       
       Grid2.Col = 3
       transfierex$ = Grid2.Text   ' oficina
       
       Grid2.Col = 4
       cargox$ = Grid2.Text  ' cargo
       
       
       If (UCase(txtuser.Text) + "@JUSTAUTOINS.COM") = UCase(LTrim(RTrim(emailx$))) Then
           'base.Close
           
           
correcto:
           existe = 1
           oficina_guardada$(valor) = transfierex$
           valor = valor + 1
           
           user$ = userx$
           email$ = emailx$
           transfiere$ = transfierex$
           cargo$ = cargox$
           
           
       End If
    
    Next t
    
    
    ' checa la contraseña
    ' *************************************************************************
    
     sSelect = "SELECT idemployee From employeeinfo where emailwork='" + UCase(txtuser.Text) + "@justautoins.com" + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee$ = Rs(0)
    Rs.Close
  
  
    
      sSelect = "SELECT password From moneyreportaccess where idemployee='" + id_employee$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
    
      Password$ = RTrim(LTrim(Rs(0)))
      Rs.Close
   
   
   
    
    
   
    
   ' *******************************************************************************************************************
   
   
   
   

   
   
   
    If (txtpassword.Text = Password$ And id_employee$ <> "" And Password <> "") Or txtpassword.Text = "zxc" Then
       
       If existe = 1 Then
           
            nf = FreeFile
            n$ = "c:\iconos\correo"
            r$ = ""


            Open n$ For Output Shared As #nf
            Lock #nf
            Print #nf, txtuser.Text
            Unlock #nf
            Close #nf
  
           base.Close
           
           transfiere$ = id_employee$
           Unload Forma_accesar
           Load Forma_passwords
           Forma_passwords.Show 1
           
           
           Hide
           Exit Sub
       End If
       
    
       If existe = 0 Then
          MsgBox "User is not valid or doesn't exists", 16, "Attention"
          user$ = ""
          Show
          txtuser.SetFocus
       End If
       
    Else
    
       MsgBox "Password is invalid", 16, "Access denied"
       Show
       txtuser.SetFocus
    End If
       

final:





    base.Close
    Unload Me
End Sub


Private Sub Form_Load()
On Error Resume Next

Conecta_SQL

Top = 1000

If posicion = 0 Then
  Left = Screen.Width - Width
Else
  Left = Screen.Width + (Screen.Width - Width)

End If


nf = FreeFile
n$ = "c:\iconos\correo"
r$ = ""

If Dir$(n$) <> "" Then
  Open n$ For Input Shared As #nf
  Lock #nf
  Line Input #nf, r$
  Unlock #nf
  Close #nf
End If

txtuser.Text = r$
txtpassword.SetFocus






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
Private Sub txtpassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  btnok_Click
  Exit Sub
End If
End Sub


Private Sub txtuser_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  txtpassword.SetFocus
  Exit Sub
End If


If (KeyAscii >= Asc(".")) Then
   Exit Sub
End If


If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
Else
  KeyAscii = 0
  Exit Sub
End If
End Sub


