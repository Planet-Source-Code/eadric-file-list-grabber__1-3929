VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "File List tool"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   6735
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "FrmListFiles.frx":0000
      Left            =   0
      List            =   "FrmListFiles.frx":0002
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Total Files"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   6735
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu NewList 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu OpenList 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Reopen 
         Caption         =   "ReOpen"
         Begin VB.Menu Clear 
            Caption         =   "Clear List"
         End
         Begin VB.Menu mnuSepMRU 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReopenSub 
            Caption         =   "None"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu LstFiles 
         Caption         =   "&With List Files"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dirty As Boolean, MRUNum As Integer

Private Sub AddToMRUX(FileNam As String)

Dim X As Integer

' x = 1 MRU Number
For X = 1 To MRUNum
    ' Checks for duplicates
    If FileNam = MRUX(X) Then Exit Sub
Next

' Opens MRU data file
Open App.Path + "\MRU.dat" For Output As #1
    ' Puts new file name if it exists
    If FileExists(FileNam) Then Print #1, FileNam
    For X = 1 To MRUNum
        ' Puts other filenames if they exist
        If X <> 15 Then If FileExists(MRUX(X)) Then Print #1, MRUX(X)
    Next
Close

' Keeps track of number of MRU
If MRUNum <> 15 Then MRUNum = MRUNum + 1

' Clears displayed MRU list
mnuReopenSub(0).Caption = "None"
mnuReopenSub(0).Enabled = False

For Num = 1 To mnuReopenSub.Count - 1
    mnuReopenSub(Num).Visible = False
Next

' Displays New MRU List
DisplayMRU

End Sub

Private Sub CreateReopenItem(ByVal menu_caption As String)

Static Menuro_Num As Integer

' Checks first MRU line to see if it's got anything (is enabled if it has anything
If mnuReopenSub(0).Enabled Then
    ' Tracking counter
    Menuro_Num = Menuro_Num + 1
    ' Loads new menu item
    Load mnuReopenSub(Menuro_Num)
Else
    ' Enables First
    mnuReopenSub(0).Enabled = True
    ' Resets counter
    Menuro_Num = 0
End If

' Puts MRU caption
mnuReopenSub(Menuro_Num).Caption = menu_caption

' Tracker
Num = Menuro_Num

End Sub

Private Sub DisplayMRU()

Dim X As Integer

' Opens MRU datafile
Open App.Path + "\MRU.dat" For Input As #1
    X = 1
    Do While Not EOF(1)
        ' Inputs MRU data
        Line Input #1, MRUX(X)
        ' Tracking counter
        X = X + 1
    Loop
Close

' Tracks MRU Number
MRUNum = X - 1

For X = 1 To MRUNum
    ' If file exists, put menu item
    If FileExists(MRUX(X)) Then CreateReopenItem (ExtractFileName(MRUX(X)))
Next

End Sub

Private Sub OpenMyList(FileNam As String)

Dim Tot As String

' Checks to see if file exists
If FileExists(FileNam) Then
    ' Clears list
    List1.Clear
    ' Gets file number
    FF = FreeFile
    ' Opens file
    Open FileNam For Input As #FF
        Do While Not EOF(FF)
            ' Gets data
            Line Input #FF, Lne
            ' Adds data to list
            List1.AddItem Lne
        Loop
    Close
    
    ' Gets total number of files
    Tot = List1.ListCount
    
    ' Displays total
    Label1.Caption = Tot + " Total Files"
    ' Adds file to MRU list
    AddToMRUX (FileNam)
Else
    ' If file doesn't exist, display warning
    MsgBox "That file does not exist", vbExclamation, "File Not Found!"
End If

End Sub

Private Sub About_Click()

' About this program
MsgBox "If you find this program or source code useful, drop me a line at (phillip@softhome.net" & vbNewLine & _
        "If you use this code, please site me in the credits.  Also tell me if you have any" & vbNewLine & _
        "suggestions or bug (fixes).", vbInformation, "About this Program"

End Sub

Private Sub Clear_Click()

Dim None As String

' Resets first MRU entry
mnuReopenSub(0).Caption = "None"
mnuReopenSub(0).Enabled = False

' Resets other MRU entries
For Num = 1 To mnuReopenSub.Count - 1
    mnuReopenSub(Num).Visible = False
Next

None = ""

' Writes blank MRU datafile
Open App.Path + "\MRU.dat" For Output As #1
Close

' Resets MRU tracking number
MRUNum = 0

End Sub

Private Sub Exit_Click()

' Unloads form
Unload Me

End Sub

Private Sub Form_Load()

'Displays MRU
DisplayMRU

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Response As String

' If list has changed it is dirty
If Dirty Then
    ' Prompts to save "dirty" list
    Response = MsgBox("Do you want to save this list", vbYesNoCancel, "List has changed!")
    ' Responds to user input
    Select Case Response
        Case vbYes
            'Save the list
            Save_Click
        Case vbNo
            ' Don't save list
            Cancel = False
        Case vbCancel
            ' Cancels quit
            Cancel = True
    End Select
End If

End Sub

Private Sub List1_DblClick()

' Removes list item
List1.RemoveItem (List1.ListIndex)

' List has changed, it is "Dirty"
Dirty = True

End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim fname As Variant, lFileSize As String

For Each fname In Data.Files
    ' Adds formatted filename and file size to list
    List1.AddItem ExtractFileName(fname + ": " + FormatFileSize(FileLen(fname)))
Next

' Indicate we did nothing with the files.
Effect = vbDropEffectNone

' List has changed, it is "Dirty"
Dirty = True

End Sub

Private Sub LstFiles_Click()

' Help with this program
MsgBox "This program lets you create lists of files by droping the files" & vbNewLine & _
        "onto the listbox.  You can search the list by typing into the" & vbNewLine & _
        "textbox, the matching files will be selected in the textbox." & vbNewLine & vbNewLine & _
        "Remove list items by double clicking on them." & vbNewLine & vbNewLine & _
        "You can open previously made lists with the MRU provided on the" & vbNewLine & _
        "File>ReOpen> Menu.  To Clear the MRU Press Clear List.", vbInformation, "Help with listfiles"
        
End Sub

Private Sub mnuReopenSub_Click(Index As Integer)

' Opens MRU Selected
OpenMyList (MRUX(Index + 1))

End Sub

Private Sub NewList_Click()

' Clears list
List1.Clear

' New lists aren't "Dirty"
Dirty = False

End Sub

Private Sub OpenList_Click()

Dim FF As Integer, FileNam As String, Lne As String, Tot As String

' Checks to see if list has been changed
If Dirty Then
    ' Prompts to save
    Response = MsgBox("Do you want to save this list", vbYesNoCancel, "List has changed!")
    ' Responds to user input
    Select Case Response
        Case vbYes
            ' Save the file
            Save_Click
        Case vbCancel
            ' Cancels the open
            Exit Sub
    End Select
End If

' Opens Common dialog box
FileNam = DialogFile(Form1, 1, "Open List", "", "TXT", App.Path, ".txt")

' Checks for valid name
If Len(FileNam) = 0 Then Exit Sub

' Opens list
OpenMyList (FileNam)

' Newly opened lists aren't "Dirty"
Dirty = False

End Sub

Private Sub Save_Click()

Dim FileNames As String, FF As Integer, Response As String, X As Integer

' Opens Common Dialog Box
FileNames = DialogFile(Form1, 2, "Save List", "", "txt", App.Path, ".TXT")

' Checks for valid name
If Len(FileNames) = 0 Then Exit Sub

' Gets file number
FF = FreeFile

' Opens list
Open FileNames For Output As #FF
    For X = 0 To List1.ListCount - 1
        ' Writes list
        Print #FF, List1.List(X)
    Next
Close

' Saved lists aren't "Dirty"
Dirty = False

' Adds saved file to MRU
AddToMRUX (FileNames)

End Sub

Private Sub Text1_Change()

' When the text changes, select the matching items
For X = 0 To (List1.ListCount - 1)
    List1.Selected(X) = False
    If Len(Text1.Text) <> 0 Then
        If Text1.Text = Left(List1.List(X), Len(Text1.Text)) Then List1.Selected(X) = True
    End If
Next

End Sub
