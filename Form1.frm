VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Entri/cari Data di ListBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()   'Tambahkan beberapa item ke 'dalam List1
   List1.AddItem "Rizky"
   List1.AddItem "Aman"
   List1.AddItem "Akhmad"
   List1.AddItem "Armanto"
   List1.AddItem "Badu"
   List1.AddItem "Bobo"
   List1.AddItem "Joko"
   List1.AddItem "Jaka"
   List1.AddItem "Parto"
   List1.AddItem "Paino"
End Sub

Private Sub Text1_Change()
  List1.ListIndex = SendMessage(List1.hwnd, _
        LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
End Sub


