VERSION 5.00
Begin VB.Form entr 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Enter Student Data"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "enter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Top             =   480
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
      Top             =   840
      Width           =   2055
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
      Top             =   1200
      Width           =   2055
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
      MaxLength       =   80
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
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
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lbl 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add a new Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1335
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
      TabIndex        =   8
      Top             =   480
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
      TabIndex        =   7
      Top             =   840
      Width           =   615
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
      TabIndex        =   6
      Top             =   1200
      Width           =   615
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
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "entr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'get the total number of entries in the ini
'add 1 to it and write the new value
Dim tmp, p As Integer, smp, wrini
tmp = GetPrivateProfileString("general", "totalnum", "", App.Path & "\data.ini")

While (p <= tmp)
smp = GetPrivateProfileString("general", "nm" & p, "", App.Path & "\data.ini")
If smp = Me.stuname.Text Then Me.Tag = smp
p = p + 1
Wend

If Me.Tag = Me.stuname Then
smp = MsgBox("The name already exists in the Database!", vbExclamation, "Error")
Exit Sub
Else
If stuname.Text <> "" And stuphone.Text <> "" And stuadd.Text <> "" And stuclass.Text <> "" Then

'create a new section [name] with the student name
'and enter all the data into it
tmp = GetPrivateProfileString("general", "totalnum", "", App.Path & "\data.ini")
tmp = tmp + 1
wrini = WriteToFile("general", "totalnum", tmp)
wrini = WriteToFile("general", "nm" & tmp, Me.stuname.Text)
wrini = WriteToFile(Me.stuname.Text, "name", Me.stuname.Text)
wrini = WriteToFile(Me.stuname.Text, "phone", Me.stuphone.Text)
wrini = WriteToFile(Me.stuname.Text, "class", Me.stuclass.Text)
wrini = WriteToFile(Me.stuname.Text, "add", Me.stuadd.Text)
Unload Me
Else
smp = MsgBox("Invalid data in fields!", vbCritical, "Error!")
Exit Sub
End If
End If
End Sub

Private Sub Form_Resize()
Me.Width = 4800
Me.Height = 3675
End Sub
