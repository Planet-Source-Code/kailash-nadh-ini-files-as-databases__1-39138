VERSION 5.00
Begin VB.Form vwdts 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Student Data"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "vwdt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox stuadd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox stuphone 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox stuclass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox stuname 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Class :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label nm 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Details about"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "vwdts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'nm.caption is the section name used for writing
Dim dg, wrini, dig As Integer
dg = MsgBox("Save Changes?", vbInformation + vbYesNo, "Save changes?")
If dg = vbYes Then
wrini = WriteToFile(Me.nm.Caption, "name", Me.stuname.Text)
wrini = WriteToFile(Me.nm.Caption, "class", Me.stuclass.Text)
wrini = WriteToFile(Me.nm.Caption, "phone", Me.stuphone.Text)
wrini = WriteToFile(Me.nm.Caption, "add", Me.stuadd.Text)
Unload Me
Else
Unload Me
End If
End Sub


Private Sub Form_Load()
'nm.caption is the section name for loading values
Me.nm.Caption = viewd.nmh.Caption
Me.stuname.Text = GetPrivateProfileString(Me.nm.Caption, "name", "", App.Path & "\data.ini")
Me.stuclass.Text = GetPrivateProfileString(Me.nm.Caption, "class", "", App.Path & "\data.ini")
Me.stuphone.Text = GetPrivateProfileString(Me.nm.Caption, "phone", "", App.Path & "\data.ini")
Me.stuadd.Text = GetPrivateProfileString(Me.nm.Caption, "add", "", App.Path & "\data.ini")
Me.Text2.Text = GetPrivateProfileString(Me.nm.Caption, "add", "", App.Path & "\data.ini")

Exit Sub
End Sub

