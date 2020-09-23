VERSION 5.00
Begin VB.Form Mainfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADDRESS BOOK"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "Mainfrm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox srpicture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      Picture         =   "Mainfrm.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Point to Contact"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton update 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   17
      ToolTipText     =   "Update Current Contact"
      Top             =   4755
      Width           =   1215
   End
   Begin VB.TextBox srname 
      DataField       =   "First"
      DataSource      =   "adrInfo"
      Height          =   285
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   13
      ToolTipText     =   "Enter first name of Contact to point to"
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton help 
      Caption         =   "HELP"
      Height          =   360
      Left            =   6720
      TabIndex        =   22
      ToolTipText     =   "Take Help on Address Book"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton delete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   16
      ToolTipText     =   "Delete Contact from Address Book"
      Top             =   4755
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1560
      Picture         =   "Mainfrm.frx":0BD4
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Prv 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      ToolTipText     =   "Previous Contact"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton nxt 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      ToolTipText     =   "Next Contact"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton lst 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      ToolTipText     =   "Last Contact"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton fst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   " First Contact"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox address3 
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   5
      ToolTipText     =   "Address"
      Top             =   2280
      Width           =   5775
   End
   Begin VB.TextBox address2 
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   4
      ToolTipText     =   "Address"
      Top             =   1920
      Width           =   5775
   End
   Begin VB.TextBox mbnum1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   9
      ToolTipText     =   "Country Code"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox phnum3 
      Height          =   285
      Left            =   5400
      MaxLength       =   8
      TabIndex        =   8
      ToolTipText     =   "Phone Number"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox phnum2 
      Height          =   285
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   7
      ToolTipText     =   "Area Code"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   23
      ToolTipText     =   "Close Address Book"
      Top             =   4755
      Width           =   1215
   End
   Begin VB.CommandButton Clear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Clear all the fields"
      Top             =   4755
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Caption         =   "ADD"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "Add contact in Address Book"
      Top             =   4755
      Width           =   1215
   End
   Begin VB.TextBox company 
      Height          =   285
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   11
      ToolTipText     =   "Company"
      Top             =   3360
      Width           =   5775
   End
   Begin VB.TextBox mail 
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   12
      ToolTipText     =   "E-mail"
      Top             =   3720
      Width           =   5775
   End
   Begin VB.TextBox address1 
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   3
      ToolTipText     =   "Address"
      Top             =   1560
      Width           =   5775
   End
   Begin VB.TextBox mbnum2 
      Height          =   285
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   10
      ToolTipText     =   "Mobile Number"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox phnum1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   6
      ToolTipText     =   "Country Code"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox lstname 
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   2
      ToolTipText     =   "Last Name"
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox fstname 
      DataField       =   "First"
      DataSource      =   "adrInfo"
      Height          =   285
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   1
      ToolTipText     =   "First Name"
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   39
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label totlabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   38
      ToolTipText     =   "Total Number of Contacts in Address Book"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label mainlabel 
      Alignment       =   2  'Center
      Caption         =   "ADDRESS BOOK "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   36
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Footer 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Label mblabel2 
      Caption         =   "Number"
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label mblabel1 
      Caption         =   "Country Code"
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label phlabel3 
      Caption         =   "Number"
      Height          =   255
      Left            =   4800
      TabIndex        =   32
      Top             =   2670
      Width           =   615
   End
   Begin VB.Label phlabel2 
      Caption         =   "Area Code"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   2670
      Width           =   855
   End
   Begin VB.Label phlabel1 
      Caption         =   "Country Code"
      Height          =   255
      Left            =   1560
      TabIndex        =   30
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label cmplabel 
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label maillabel 
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Adrlabel 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label mblabel 
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label phlabel 
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lstlabel 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label fstlabel 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mailerror As String
Dim hwndhelp As Long
Private Sub address1_GotFocus()
Footer.Caption = "First Line of address of Contact"
End Sub

Private Sub address2_GotFocus()
Footer.Caption = "Second Line of address of Contact"
End Sub

Private Sub address3_GotFocus()
Footer.Caption = "Third Line of address of Contact"
End Sub

Private Sub delete_Click()
On Error GoTo error
If rs.RecordCount = 0 Then
MsgBox "No Contacts in Address Book to Delete.", vbOKOnly + vbCritical, "Address Book - Error"
Exit Sub
End If

Value& = MsgBox("Do you really want to delete this Contact from Address Book ?", vbYesNo + vbInformation, "Address Book - Delete Confirm")
If Value& = vbYes Then
rs.delete
MsgBox "Contact deleted from Address Book.", vbOKOnly + vbInformation, "Address Book - Contact Deleted"
totlabel.Caption = rs.RecordCount
If rs.RecordCount > 0 Then
Call nxt_Click
Else
Call Clear_Click
End If
End If
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End

End Sub

Private Sub delete_GotFocus()
Footer.Caption = "Delete Current Contact from Address Book"
End Sub
Private Sub exit_Click()
Unload Me
End Sub

Private Sub exit_GotFocus()
Footer.Caption = "Close Address Book"
End Sub

Private Sub Clear_Click()
On Error GoTo error
code = "add"
Call clrcolor
If rs.RecordCount > 0 Then
rs.MoveFirst
End If
filests = "Clear"
fstname.Text = ""
lstname.Text = ""
phnum1.Text = ""
phnum2.Text = ""
phnum3.Text = ""
mbnum1.Text = ""
mbnum2.Text = ""
address1.Text = ""
address2.Text = ""
address3.Text = ""
company.Text = ""
mail.Text = ""
srname.Text = ""
fstname.SetFocus
update.Enabled = False
delete.Enabled = False
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub

Private Sub Clear_GotFocus()
Footer.Caption = "Clear all the input capable fields"
End Sub

Private Sub company_GotFocus()
Footer.Caption = "Company of Contact"
End Sub

Private Sub Form_Initialize()
If App.PrevInstance Then
        MsgBox "Another copy of the Address Book is already open.", vbCritical + vbOKOnly, "Address Book - Open Error"
        OpenError = True
        Unload Me
        Set Mainfrm = Nothing
        End
End If
App.TaskVisible = False
End Sub

Private Sub Form_Load()
On Error GoTo error
initdb
code = "add"
If rs.RecordCount > 0 Then
rs.MoveFirst
filests = "First"
fill
Else
MsgBox "Address Book is empty. Click ADD button to add contacts.", vbOKOnly + vbInformation, "Address Book - Empty"
update.Enabled = False
delete.Enabled = False
End If
totlabel.Caption = rs.RecordCount
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub

Private Sub fst_Click()
On Error GoTo error
srname.Text = ""
If rs.RecordCount = 0 Then
MsgBox "Address Book is empty.", vbOKOnly + vbInformation, "Address Book - Empty"
Exit Sub
End If
If filests = "First" Then
    MsgBox "You are at FIRST Contact of Address Book.", vbOKOnly + vbInformation, "Address Book - Information"
Exit Sub
End If
rs.MoveFirst
If rs.BOF Or filests = "First" Then
            MsgBox "You are at FIRST Contact of Address Book.", vbOKOnly + vbInformation, "Address Book - Information"
            rs.MoveFirst
            Exit Sub
        Else
            rs.Edit
            fill
        End If
filests = "First"
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub

Private Sub fst_GotFocus()
Footer.Caption = "Show First Contact in Address Book"
End Sub

Private Sub fstname_GotFocus()
Footer.Caption = "First Name of Contact"
End Sub

Private Sub help_Click()
ShowHelpFile (App.Path & "\AddBookHelp.chm")
If hwndhelp = 0 Then
MsgBox "An error has occurred while opening the help file for Address Book" & vbCrLf & _
"Or the help file 'AddBookHelp.chm' not found in the current path." & vbCrLf & _
"Please make sure that the help file 'AddBookHelp.chm' is there is the current path.", vbOKOnly + vbCritical, "Address Book - Error"
End If
End Sub

Private Sub help_GotFocus()
Footer.Caption = "Address Book Help"
End Sub
Private Sub lst_Click()
On Error GoTo error
srname.Text = ""
If rs.RecordCount = 0 Then
MsgBox "Address Book is empty.", vbOKOnly + vbInformation, "Address Book - Empty"
Exit Sub
End If

If filests = "Last" Then
   MsgBox "You are at LAST Contact of Address Book.", vbOKOnly + vbInformation, "Address Book - Information"
Exit Sub
End If

rs.MoveLast
If rs.EOF Or filests = "Last" Then
            MsgBox "You are at LAST Contact of Address Book.", vbOKOnly + vbInformation, "Address Book - Information"
            rs.MoveLast
            Exit Sub
        Else
            rs.Edit
            fill
        End If
filests = "Last"
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub

Private Sub lst_GotFocus()
Footer.Caption = "Show Last Contact in Address Book"
End Sub

Private Sub lstname_GotFocus()
Footer.Caption = "Last Name of Contact"
End Sub

Private Sub mail_GotFocus()
Footer.Caption = "Internet E-mail address of Contact"
End Sub

Private Sub mbnum1_GotFocus()
Footer.Caption = "Country Code for Mobile Number"
End Sub

Private Sub mbnum2_GotFocus()
Footer.Caption = "Mobile Number of Contact"
End Sub

Private Sub nxt_Click()
On Error GoTo error
srname.Text = ""
If rs.RecordCount = 0 Then
MsgBox "Address Book is empty.", vbOKOnly + vbInformation, "Address Book - Empty"
Exit Sub
End If

If filests = "Clear" Then
    filests = "First"
    rs.Edit
    fill
Else
    rs.MoveNext
    If rs.EOF Then
          MsgBox "You are at LAST Contact of Address Book.", vbOKOnly + vbInformation, "Address Book - Information"
          filests = "Last"
          rs.MoveLast
          Exit Sub
    Else
          filests = "Middle"
          rs.Edit
          fill
    End If
End If
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub

Private Sub nxt_GotFocus()
Footer.Caption = "Show Next Contact in Address Book"
End Sub

Private Sub add_Click()
On Error GoTo error
Call clrcolor

Footer.Caption = "Validating user input"

If fstname.Text = "" Then
fstname.BackColor = &HFF00&
MsgBox "First Name should be filled.It is mandatory.", vbOKOnly + vbCritical, "Address Book - Error"
fstname.SetFocus
Exit Sub
End If

If lstname.Text = "" Then
lstname.BackColor = &HFF00&
MsgBox "Last Name should be filled.It is mandatory.", vbOKOnly + vbCritical, "Address Book - Error"
lstname.SetFocus
Exit Sub
End If

If phnum1 <> "" Then
If phnum2 = "" Then
phnum2.BackColor = &HFF00&
MsgBox "Area Code must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
phnum2.SetFocus
Exit Sub
End If
End If

If phnum1 <> "" Then
If phnum3 = "" Then
phnum3.BackColor = &HFF00&
MsgBox "Phone Number must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
phnum3.SetFocus
Exit Sub
End If
End If

If phnum2 <> "" Then
If phnum1 = "" Then
phnum1.BackColor = &HFF00&
MsgBox "Country Code must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
phnum1.SetFocus
Exit Sub
End If
End If

If phnum2 <> "" Then
If phnum3 = "" Then
phnum3.BackColor = &HFF00&
MsgBox "Country Code must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
phnum3.SetFocus
Exit Sub
End If
End If

If phnum3 <> "" Then
If phnum1 = "" Then
phnum1.BackColor = &HFF00&
MsgBox "Country Code must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
phnum1.SetFocus
Exit Sub
End If
End If

If phnum3 <> "" Then
If phnum2 = "" Then
phnum1.BackColor = &HFF00&
MsgBox "Country Code must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
phnum1.SetFocus
Exit Sub
End If
End If

If IsNumeric(phnum1) = False And phnum1 <> "" Then
phnum1.BackColor = &HFF00&
MsgBox "Country Code entered is not numeric.", vbOKOnly + vbCritical, "Address Book - Error"
phnum1.SetFocus
Exit Sub
End If

If IsNumeric(phnum2) = False And phnum2 <> "" Then
phnum2.BackColor = &HFF00&
MsgBox "Area Code entered is not numeric.", vbOKOnly + vbCritical, "Address Book - Error"
phnum2.SetFocus
Exit Sub
End If

If IsNumeric(phnum3) = False And phnum3 <> "" Then
phnum3.BackColor = &HFF00&
MsgBox "Phone Number entered is not numeric.", vbOKOnly + vbCritical, "Address Book - Error"
phnum3.SetFocus
Exit Sub
End If


If mbnum1 <> "" Then
If mbnum2 = "" Then
mbnum2.BackColor = &HFF00&
MsgBox "If Conutry Code filled, Mobile Number must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
mbnum2.SetFocus
Exit Sub
End If
End If

If mbnum2 <> "" Then
If mbnum1 = "" Then
mbnum1.BackColor = &HFF00&
MsgBox "If Mobile Number filled, Country Code must also be filled.", vbOKOnly + vbCritical, "Address Book - Error"
mbnum1.SetFocus
Exit Sub
End If
End If

If IsNumeric(mbnum1) = False And mbnum1 <> "" Then
mbnum1.BackColor = &HFF00&
MsgBox "Country Code entered is not numeric.", vbOKOnly + vbCritical, "Address Book - Error"
mbnum1.SetFocus
Exit Sub
End If

If IsNumeric(mbnum2) = False And mbnum1 <> "" Then
mbnum2.BackColor = &HFF00&
MsgBox "Mobile Number entered is not numeric.", vbOKOnly + vbCritical, "Address Book - Error"
mbnum2.SetFocus
Exit Sub
End If

If address1.Text = "" And address2.Text = "" And address3.Text = "" Then
address1.BackColor = &HFF00&
MsgBox "Address should be filled.It is mandatory.", vbOKOnly + vbCritical, "Address Book - Error"
address1.SetFocus
Exit Sub
End If

If mail.Text <> "" Then
mailerror = checkMailVal(mail.Text)
If mailerror <> "" Then
mail.BackColor = &HFF00&
MsgBox mailerror, vbOKOnly + vbCritical, "Address Book - Error"
mail.SetFocus
Exit Sub
End If
End If

Select Case (code)

Case "add"

Set rs = db.OpenRecordset("SELECT * FROM Contacts WHERE First LIKE '" & fstname.Text & "'" & "")

If rs.RecordCount <> 0 Then

If MsgBox("A contact with first name '" & UCase(fstname.Text) & "' already exist. Do you want to add this one also ?", vbYesNo + vbInformation, "Address Book - Already Exist") = vbYes Then
rs.AddNew
rs("First") = UCase(fstname.Text)
rs("Last") = UCase(lstname.Text)
rs("Ctcd1") = phnum1.Text
rs("Arcd") = phnum2.Text
rs("Phone") = phnum3.Text
rs("Ctcd2") = mbnum1.Text
rs("Mobile") = mbnum2.Text
rs("Address1") = address1.Text
rs("Address2") = address2.Text
rs("Address3") = address3.Text
rs("Company") = company.Text
rs("Email") = mail.Text
rs.update
totlabel.Caption = totlabel.Caption + 1
MsgBox "Record successfully added to Address Book", vbOKOnly + vbInformation, "Address Book - Contact Added"
Call Clear_Click
Else
MsgBox "Record not added to Address Book", vbOKOnly + vbInformation, "Address Book - Contact Not Added"
End If

Else
rs.AddNew
rs("First") = UCase(fstname.Text)
rs("Last") = UCase(lstname.Text)
rs("Ctcd1") = phnum1.Text
rs("Arcd") = phnum2.Text
rs("Phone") = phnum3.Text
rs("Ctcd2") = mbnum1.Text
rs("Mobile") = mbnum2.Text
rs("Address1") = address1.Text
rs("Address2") = address2.Text
rs("Address3") = address3.Text
rs("Company") = company.Text
rs("Email") = mail.Text
rs.update
totlabel.Caption = totlabel.Caption + 1
MsgBox "Record successfully added to Address Book", vbOKOnly + vbInformation, "Address Book - Contact Added"
Call Clear_Click

End If

Case "update"
Set rs = db.OpenRecordset("SELECT * FROM Contacts WHERE First LIKE '" & UCase(fstname.Text) & "'" & "")
rs.Edit
rs("First") = UCase(fstname.Text)
rs("Last") = UCase(lstname.Text)
rs("Ctcd1") = phnum1.Text
rs("Arcd") = phnum2.Text
rs("Phone") = phnum3.Text
rs("Ctcd2") = mbnum1.Text
rs("Mobile") = mbnum2.Text
rs("Address1") = address1.Text
rs("Address2") = address2.Text
rs("Address3") = address3.Text
rs("Company") = company.Text
rs("Email") = mail.Text
rs.update
MsgBox "Record successfully updated in Address Book", vbOKOnly + vbInformation, "Address Book - Contact Updated"
code = "add"
End Select
Set rs = db.OpenRecordset("SELECT * FROM Contacts")
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub
Private Sub clrcolor()
fstname.BackColor = &H80000005
lstname.BackColor = &H80000005
phnum1.BackColor = &H80000005
phnum2.BackColor = &H80000005
phnum3.BackColor = &H80000005
mbnum1.BackColor = &H80000005
mbnum2.BackColor = &H80000005
address1.BackColor = &H80000005
address2.BackColor = &H80000005
address3.BackColor = &H80000005
company.BackColor = &H80000005
mail.BackColor = &H80000005
End Sub

Private Sub add_GotFocus()
Footer.Caption = "Add to the Address Book"
End Sub

Private Sub phnum1_GotFocus()
Footer.Caption = "Country Code for Phone Number"
End Sub

Private Sub phnum2_GotFocus()
Footer.Caption = "Area Code for Phone Number"
End Sub

Private Sub phnum3_GotFocus()
Footer.Caption = "Phone Number of Contact"
End Sub

Private Sub Prv_Click()
On Error GoTo error
srname.Text = ""
If rs.RecordCount = 0 Then
MsgBox "Address Book is empty.", vbOKOnly + vbInformation, "Address Book - Empty"
Exit Sub
End If

If filests = "Clear" Then
filests = "First"
            rs.Edit
            fill
Else
rs.MovePrevious
        If rs.BOF Then
            MsgBox "You are at FIRST Contact of Address Book.", vbOKOnly + vbInformation, "Address Book - Information"
            filests = "First"
            rs.MoveFirst
            Exit Sub
        Else
            filests = "Middle"
            rs.Edit
            fill
        End If
End If
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub


Private Sub Prv_GotFocus()
Footer.Caption = "Show Previous Contact in Address Book"
End Sub

Private Function fill()
delete.Enabled = True
update.Enabled = True
fstname.Text = rs("First")
lstname.Text = rs("Last")
phnum1.Text = rs("Ctcd1")
phnum2.Text = rs("Arcd")
phnum3.Text = rs("Phone")
mbnum1.Text = rs("Ctcd2")
mbnum2.Text = rs("Mobile")
address1.Text = rs("Address1")
address2.Text = rs("Address2")
address3.Text = rs("Address3")
company.Text = rs("Company")
mail.Text = rs("Email")
End Function

Private Sub srname_Change()
On Error GoTo error
If srname.Text <> "" Then
Set rs = db.OpenRecordset("SELECT * FROM Contacts WHERE First LIKE '" & UCase(srname.Text) & "'" & "& '*'")
If rs.RecordCount > 0 Then
fill
End If
End If
Set rs = db.OpenRecordset("SELECT * FROM Contacts")
filests = "Middle"
Exit Sub
error:
MsgBox "Some Unexpected Error has occured. Application will close now.", vbCritical + vbOKOnly, "Address Book - Unexpected Error"
End
End Sub

Private Sub srname_GotFocus()
Footer.Caption = "Point to Contact in Address Book"
End Sub

Private Sub update_Click()
code = "update"
Value& = MsgBox("Are you sure to UPDATE this Contact in Address Book ?.", vbYesNo + vbInformation, "Address Book - Delete Confirm")
If Value& = vbYes Then
Call add_Click
End If
End Sub

Private Sub update_GotFocus()
Footer.Caption = "Update Current Contact in Address Book"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("This will close the Address Book." & vbCrLf & vbCrLf & _
"           Are you sure?", vbYesNo + vbQuestion, "Address Book - Close") = vbYes Then
If OpenError Then Exit Sub
Value& = MsgBox("Would you like Address Book to run at startup?", vbYesNo + vbQuestion, "Address Book - Startup Run")
    If Value& = vbYes Then
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Address Book", App.Path & "\" & App.EXEName & ".exe"
            Else
        SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", "Address Book", "<NonRun>"
    End If
db.Close
Set Mainfrm = Nothing
End
Else
Cancel = True
End If
End Sub
'Opens the compiled help file
Private Function ShowHelpFile(strFilename As String) As Long
    'The return value is the window handle o
    '     f the created help window.
    hwndhelp = HtmlHelp(hWnd, strFilename, HH_DISPLAY_TOPIC, 0)
    
End Function
