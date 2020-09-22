VERSION 5.00
Begin VB.Form frmSell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "frmSell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3465
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.HScrollBar scr 
      Height          =   135
      Left            =   120
      Min             =   1
      TabIndex        =   1
      Top             =   840
      Value           =   1
      Width           =   855
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "&Sell"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblInform 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter how many would you like to sell"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item(s)"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   525
      Width           =   465
   End
End
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmSell
End Sub

Private Sub cmdSell_Click()
    Dim i As Integer 'Counter
    Dim j As Integer 'Store Value
    For i = 1 To 17
        'Find the array number of the foods you want to buy and store it in j
        If Foods(i) = frmMain.lstItems.SelectedItem Then j = i
    Next
    'Work out how much money you gain from the sell
    Credit = Credit + ((frmMain.lstFoods.ListItems(j).ListSubItems(1) - Avg(frmMain.lstItems.SelectedItem.Index)) * txtQty.text) + (Avg(frmMain.lstItems.SelectedItem.Index) * txtQty.text)
    frmMain.lblCash.Caption = IIf(Credit <> 0, Format(Credit, "###,###,###"), 0)
    'Reduce the amount of units you have according to how many units you sold
    Quantity(j) = Quantity(j) - txtQty.text
    'If you sold all your units the average is now 0 since you have no more units
    If Quantity(j) = 0 Then Avg(j) = 0
    'Remove the item from your items list since you just sold it
    If Quantity(j) = 0 Then frmMain.lstItems.ListItems.Remove (frmMain.lstItems.SelectedItem.Index)
    'If you didn't sell all your units update how many you have left
    If Not Quantity(j) = 0 Then frmMain.lstItems.SelectedItem.ListSubItems(2) = Quantity(j)
    'Update amount of space
    iSpace = iSpace + txtQty.text
    'reduce the amount of space occupied
    Used = Used - txtQty.text
    frmMain.lblItems = "Items: " & Used & " of " & TotalSpace
    Sold = True
    PlaySound SDir & "cashreg.wav", 0, 3
    Unload frmSell
End Sub

Private Sub Form_Load()
    scr.Max = frmMain.lstItems.SelectedItem.ListSubItems(2)
    scr.Value = frmMain.lstItems.SelectedItem.ListSubItems(2)
    txtQty.SelLength = Len(txtQty)
End Sub

Private Sub scr_Change()
    txtQty.text = scr.Value
End Sub

