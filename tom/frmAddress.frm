VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Control Example"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">>"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset


Private Sub cmdAdd_Click()
newName = InputBox("Please Enter The Name To Add", "AddressBook")
newEmail = InputBox("Please Enter The Email To Add", "AddressBook")
newPhone = InputBox("Please Enter The Phone Number To Add", "AddressBook")
newFax = InputBox("Please Enter The Fax Number To Add", "AddressBook")

With rs
.AddNew
!Name = LCase(newName)
!Email = LCase(newEmail)
!Phone = LCase(newPhone)
!Fax = LCase(newFax)
.Update
End With

MsgBox (newName & " Was Added To The Address Book")


End Sub

Private Sub cmdBack_Click()
On Error Resume Next
rs.MovePrevious
txtName = rs.Fields("Name")
txtEmail = rs.Fields("Email")
txtPhone = rs.Fields("Phone")
txtFax = rs.Fields("Fax")
End Sub


Private Sub cmdDelete_Click()

Set rs = db.OpenRecordset("SELECT ppl.name, ppl.email, ppl.phone, ppl.fax From ppl WHERE ppl.name = " + Chr$(34) + txtName.Text + Chr$(34) + ";")


rs.MoveFirst
Do Until rs.EOF
rs.Delete
rs.MoveNext
Loop
db.Close




End Sub

Private Sub cmdForward_Click()
On Error Resume Next
rs.MoveNext
txtName = rs.Fields("Name")
txtEmail = rs.Fields("Email")
txtPhone = rs.Fields("Phone")
txtFax = rs.Fields("Fax")
End Sub

Private Sub cmdSearch_Click()
Set rs = db.OpenRecordset("SELECT * FROM ppl")
NameQuery = InputBox("Enter A Name To Search For", "Name Query")


rs.MoveFirst
Do Until rs.EOF

If rs.Fields("name") Like "*" & LCase(NameQuery) & "*" Then
txtName = rs.Fields("Name")
txtEmail = rs.Fields("Email")
txtPhone = rs.Fields("Phone")
txtFax = rs.Fields("Fax")
Exit Sub
Else
rs.MoveNext
End If


Loop







End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path + "/ppl.mdb")
Set rs = db.OpenRecordset("ppl")

If rs.EOF Then
MsgBox "Please Add Someone To The Databse"
Else
rs.MoveFirst
txtName = rs.Fields("Name")
txtEmail = rs.Fields("Email")
txtPhone = rs.Fields("Phone")
txtFax = rs.Fields("Fax")
End If

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ToolTipText = txtName.Text
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ToolTipText = txtEmail.Text
End Sub
