VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   2640
   ClientTop       =   1620
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3030
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Check Spellings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1410
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3030
      Width           =   1665
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2955
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5212
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // Mehdi - mehdi240684@yahoo.com      //
' // Required References :              //
' // Microsoft Word 10.0 Object Library //

Public WordApp              As Word.Application
Public xRange               As Range
Public CorrectionCol        As SpellingSuggestions
Public SpellCol             As ProofreadingErrors
Private WordCount           As Integer

Private Sub Command1_Click()
    CheckMehdi
End Sub

Private Sub Command2_Click()
On Error GoTo x
    WordApp.Quit 0
x:
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo x
    WordApp.Quit 0
    Exit Sub
x:
    End
End Sub



Public Sub CheckMehdi()

On Error GoTo WordError
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False
    WordApp.Documents.Add
    Set xRange = WordApp.ActiveDocument.Range
    xRange.InsertAfter Text1.Text
    Set SpellCol = xRange.SpellingErrors
    If SpellCol.Count > 0 Then
        frmErrorList.List1.Clear: frmErrorList.List2.Clear
        For WordCount = 1 To SpellCol.Count
            frmErrorList.List1.AddItem SpellCol.Item(WordCount)
        Next
    End If
    If frmErrorList.List1.ListCount = 0 Then
        MsgBox "No Spelling Errors Found!", vbInformation, "Information"
    Else
        frmErrorList.Show vbModal
    End If
    Exit Sub
WordError:
    MsgBox "ERROR: " & Err.Number & Chr(13) & Err.Description

End Sub
