VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   915
      Left            =   1725
      TabIndex        =   2
      Top             =   2430
      Width           =   3150
   End
   Begin Project1.wbxSheetbar wbxSheetbar1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.wbxProgress wbxProgress1 
      Height          =   330
      Left            =   570
      TabIndex        =   0
      Top             =   1545
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   582
      ShowValue       =   0   'False
      ProgressModel   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   360
      Left            =   1740
      TabIndex        =   3
      Top             =   660
      Width           =   2865
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For ic = 0 To 1000
wbxProgress1.Value = ic \ 10
Refresh
Next
End Sub

Private Sub Form_Load()
wbxSheetbar1.AddSheet "test1", 1
wbxSheetbar1.AddSheet "test2", 2
wbxSheetbar1.AddSheet "test3", 3
wbxSheetbar1.AddSheet "test4", 4
End Sub

Private Sub wbxSheetbar1_SheetClick(Index As Long)
Label1 = wbxSheetbar1.SheetCaption(Index) & " selected!"
End Sub
