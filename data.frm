VERSION 5.00
Begin VB.Form viewd 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Entries in DB"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "data.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Hidden button"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text5 
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
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Enter student's Name."
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox nm 
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
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Enter student's Name."
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label nmh 
      Height          =   135
      Left            =   4080
      TabIndex        =   5
      Top             =   3840
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4575
      X2              =   120
      Y1              =   720
      Y2              =   735
   End
End
Attribute VB_Name = "viewd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################# INI Files as Databases ##########
'Code by Kailash Nadh , 15 yrs, India
'Please visit http://kbn.rom.cd for cool softwares
'all below 500 KB! (3D Clock, talking calculator, pass detector and more..)
'
'If you have any comments, please mailto bal_kumar@satyam.net.in

'An advanced program using INI as databasescan be downloaded from
'my site. Its Quiz Master, 180 KB, no source code!
'A must see one!!

'################ IMPORTANT #############
'YOU CAN GIVE ANY EXTENSION TO THE INI FILE, NO NEED OF INI
'IT CAN BE .DAT, .DLL , .DDD OR ANYTHING!!!


Private Sub Command1_Click()
Dim dts, msgb, pk As Integer, gt, pkp
pk = 1 'an integer for while loop
'Get the total number of entries from the ini file
dts = GetPrivateProfileString("general", "totalnum", "", App.Path & "\data.ini")

'do the loop from 1 to total & each time load the key "nm" & the number
'(the increment integer pk). If the value nm+(the number) matcher
'the search item, then display a message, else notfound message
While (pk <= dts)
gt = GetPrivateProfileString("general", "nm" & pk, "", App.Path & "\data.ini")
If gt = Me.nm.Text Then pkp = gt
pk = pk + 1
Wend

If Me.nm.Text <> "" And pkp <> "" Then
Call fnd 'function for displaying the details if search was successfull
Else
msgb = MsgBox("No matching Entries were found.", vbExclamation, "Error!")
End If
Exit Sub
End Sub

Function fnd()
Dim msgb2
msgb2 = MsgBox("An Entry was found." & vbCrLf & "View Details?", vbInformation + vbYesNo, "Entry Found!")
If msgb2 = vbYes Then
Me.nmh.Caption = Me.nm.Text
vwdts.Show vbModal, Me
End If
If msgb2 = vbNo Then msgb2 = ""
End Function


Private Sub Command2_Click()
'view details button
If Me.lst.Text <> "" Then
Me.nmh.Caption = Me.lst.Text
vwdts.Show vbModal, Me
End If
End Sub


Private Sub Command3_Click()
'delete an entry
Dim ty, ber, kai, tst
tst = MsgBox("Are you sure want to delete the entry?", vbExclamation + vbYesNo, "Delete Entry?")

If tst = vbYes Then
If Me.Text5.Text <> "" Then
'clear the deleted entry from the listbox
ty = Me.lst.ListCount
For ber = 0 To ty - 1

If ber = Me.Text5.Text Then
kai = Me.Text5.Text
Me.lst.RemoveItem kai
Text5.Text = ""
'no need for this button, but i just gave!
Call Command4_Click
End If

Next ber

End If
End If

Exit Sub
End Sub

Private Sub Command4_Click()
'this button code deletes the values from the ini file
'Command4.tag is [section] in the inifile.
'Within that section, delete all the details
'The [section] remains there without any data.
'So it won't be taken care of!
Me.lst.Refresh
Call DeleteFromFile(Me.Command4.Tag, "name")
Call DeleteFromFile(Me.Command4.Tag, "class")
Call DeleteFromFile(Me.Command4.Tag, "phone")
Call DeleteFromFile(Me.Command4.Tag, "add")

Dim wrini, d As Integer, rtval, wrinis, ft As Integer, po As String, pos, bmw
rtval = GetPrivateProfileString("general", "totalnum", "", App.Path & "\data.ini")

'delete the nm+number value from the [general] section
'eg: delete, nm6=Kailash
ft = 1
d = 1
While (ft <= rtval)
Call DeleteFromFile("general", "nm" & ft)
ft = ft + 1
Wend

'rewrite all the data in order so that that
'the values come in order. eg:
'if nm3 is deleted, remaining is nm1,nm2,nm4
'Rewrite it into nm1,nm2,nm3
Me.lst.Refresh
po = 1
Dim b, mt
b = Me.lst.ListCount
For mt = 0 To b - 1
Me.lst.ListIndex = mt
Call WriteToFile("general", "nm" & po, Me.lst.Text)
po = po + 1
Next mt

'now get the total and write it into the ini
rtval = rtval - 1
If rtval > 1 Then
wrini = WriteToFile("general", "totalnum", rtval)
Else
wrini = WriteToFile("general", "totalnum", 0)
End If

Exit Sub
End Sub

Private Sub Form_Load()
'load names from the [general] section in order
'with a loop from 1 to total to the list box
Dim gnrl, a As Integer, nm1
a = 1
gnrl = GetPrivateProfileString("general", "totalnum", "", App.Path & "\data.ini")

While (a <= gnrl)
nm1 = GetPrivateProfileString("general", "nm" & a, "", App.Path & "\data.ini")
If nm1 <> "" Then
Me.lst.AddItem nm1
End If
a = a + 1
Wend

Exit Sub
End Sub

Private Sub lst_Click()
'set lsist box's tag to its listindex.
'this may be needed for deleting a value from the list box
Me.Text5.Text = Me.lst.ListIndex
Me.Command4.Tag = Me.lst.Text
End Sub

   
Private Sub Form_Resize()
'keep the form compact and small!
Me.Width = 4860
Me.Height = 3705
End Sub
