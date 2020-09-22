VERSION 5.00
Begin VB.Form frmBuy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buy"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "frmBuy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3525
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.HScrollBar scr 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item(s)"
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   525
      Width           =   465
   End
   Begin VB.Label lblInform 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter how many would you like to purchase"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "frmBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuy_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    j = 1
    'If cash < than the amount of money this will cost don't go ahead
    If Credit < frmMain.lstFoods.SelectedItem.ListSubItems(1).text * txtQty.text Then Exit Sub
    frmMain.lstItems.ListItems.Clear
    Credit = Credit - frmMain.lstFoods.SelectedItem.ListSubItems(1).text * txtQty.text 'Update cash remaining
    frmMain.lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
    If Quantity(frmMain.lstFoods.SelectedItem.Index) > 0 Then 'if selected foods has been purchased before
        Quantity(frmMain.lstFoods.SelectedItem.Index) = Quantity(frmMain.lstFoods.SelectedItem.Index) + txtQty.text 'Add bought units onto already got units
        Else
        Quantity(frmMain.lstFoods.SelectedItem.Index) = txtQty.text 'Haven't been bought before so just add quantity bought
    End If
    If Avg(frmMain.lstFoods.SelectedItem.Index) > 0 Or Not Quantity(frmMain.lstFoods.SelectedItem.Index) = 0 Then  'If You have already bought the item or if you have stolen
        Avg(frmMain.lstFoods.SelectedItem.Index) = ((Quantity(frmMain.lstFoods.SelectedItem.Index) - txtQty.text) * (Avg(frmMain.lstFoods.SelectedItem.Index)) + txtQty.text * frmMain.lstFoods.SelectedItem.ListSubItems(1)) / (txtQty.text + (Quantity(frmMain.lstFoods.SelectedItem.Index) - txtQty.text)) 'Get Avererage.
        Else
        Avg(frmMain.lstFoods.SelectedItem.Index) = frmMain.lstFoods.SelectedItem.ListSubItems(1).text 'It hasn't been bought before so just set the average as the current price
    End If
    'Avg is a double variable because if the user bought
    '180 units at 40 cash and then bought one by one until they reached 20 at 50 cash
    'the average would remain 40 instead of 41 or 42 or something becuase  Long
    'Integers don't contain decimals and if the result returned was 40.25 it would
    'round it to 40 so no matter how many units you went up by one it would remain 40
    'unless you increased by a higher number in one turn like 30 or 40, because
    'even if there is still decimal places the number would change the average
    'by several whole numbers. With Double or Single variables if it recieved 40.25
    'it would keep the value 40.25 and below express it rounded off but actually keep
    'the decimal stored so if you did it again and it got 0.40 it would add on to
    'that 40.25 making it 40.65 and then rounding it off to 41 where if it was a long
    'it would have made 40.40 and rounded it off to 40 because the previous decimal
    'wasn't stored. I hope you understand anyway.
    For i = 1 To 17
        If Quantity(i) > 0 Then 'Add all foods from 1 to 17 that have been purchased (Quantity is how many purchased, if quantity = 0 none have been purchased)
            frmMain.lstItems.ListItems.Add j, , frmMain.lstFoods.ListItems(i)
            frmMain.lstItems.ListItems(j).ListSubItems.Add , , Round(Avg(i), 0)
            frmMain.lstItems.ListItems(j).ListSubItems.Add , , Quantity(i)
            j = j + 1 'Done, now goto next j
        End If
    Next
    iSpace = iSpace - txtQty.text 'iSpace left from 200 spaces (standard)
    Used = Used + txtQty.text 'How many spaces used
    frmMain.lblItems = "Items: " & Used & " of " & TotalSpace 'Update label
    PlaySound SDir & "cashreg.wav", 0, 3
    Unload frmBuy
End Sub

Private Sub cmdCancel_Click()
    Unload frmBuy
End Sub

Private Sub Form_Load()
    scr.Min = 1
    If Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text) > iSpace Then
        scr.Max = iSpace
        scr.Value = iSpace
        txtQty.SelLength = Len(txtQty)
        Exit Sub
    End If
    scr.Max = Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text)
    scr.Value = Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text)
    txtQty.text = Int(Credit / frmMain.lstFoods.SelectedItem.ListSubItems(1).text)
    txtQty.SelLength = Len(txtQty)
End Sub

Private Sub scr_Change()
    txtQty.text = scr.Value
End Sub

