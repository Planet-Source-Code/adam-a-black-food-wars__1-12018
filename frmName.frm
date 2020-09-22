VERSION 5.00
Begin VB.Form frmName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter your name"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2430
   Icon            =   "frmName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your name"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    If Len(txtName) > HSL.GetNameLength Then
        MsgBox "Names can only be up to " & HSL.GetNameLength & " characters long.", vbInformation, "Name"
        Exit Sub
    ElseIf Len(Trim(txtName)) = 0 Then
        MsgBox "Invalid name.", vbCritical, "High Scores"
        Exit Sub
    End If
    SaveSetting "Food Wars", "Startup", "Name", txtName.text
    User = txtName.text
    Load frmMain
    frmMain.Show
    Unload frmName
End Sub

Private Sub Form_Load()
    Dim Usr As String
    Dim iLen As Long
    Usr = Space$(255)
    GetUserName Usr, 255
    HSL.SetNameLength 15
    txtName.text = GetSetting("Food Wars", "Startup", "Name", Usr)
    txtName.SelLength = Len(txtName.text)
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Trim(txtName.text) <> "" Then
        If Len(txtName) > HighScores.GetNameLength Then
            MsgBox "Names can only be up to " & HighScores.GetNameLength & " characters long.", vbInformation, "Name"
            Exit Sub
        ElseIf Len(Trim(txtName)) = 0 Then
            MsgBox "Invalid name.", vbCritical, "High Scores"
            Exit Sub
        End If
        User = txtName.text
        Load frmMain
        frmMain.Show
        Unload frmName
    End If
End Sub
