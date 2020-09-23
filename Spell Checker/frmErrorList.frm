VERSION 5.00
Begin VB.Form frmErrorList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spelling Errors"
   ClientHeight    =   3075
   ClientLeft      =   4350
   ClientTop       =   2745
   ClientWidth     =   4365
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4245
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1995
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   2130
         TabIndex        =   1
         ToolTipText     =   "Double Click Here To Replace Errors With Suggestions!"
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suggesstions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   90
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Errors"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   3
         Top             =   90
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmErrorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CheckPos As Single

Private Sub Form_Load()
    CheckPos = 0
End Sub

Private Sub List1_Click()
Screen.MousePointer = vbHourglass
Set Form1.CorrectionCol = Form1.WordApp.GetSpellingSuggestions(Form1.SpellCol.Item(List1.ListIndex + 1))
List2.Clear
For a = 1 To Form1.CorrectionCol.Count
    List2.AddItem Form1.CorrectionCol.Item(a)
Next
Screen.MousePointer = vbDefault
End Sub

Private Sub List2_DblClick()
    If InStr(Form1.Text1.Text, List1.Text) = 0 Then Exit Sub
    Form1.Text1.Find List1.Text, CheckPos, Len(Form1.Text1.Text)
    Form1.Text1.SelText = List2.Text
    CheckPos = (Form1.Text1.SelStart)
    'List1.RemoveItem (List1.ListIndex)
    List1.Clear
    List2.Clear
    Unload Me
    Form1.WordApp.Quit 0
    Form1.CheckMehdi
    'Unload Me
End Sub
