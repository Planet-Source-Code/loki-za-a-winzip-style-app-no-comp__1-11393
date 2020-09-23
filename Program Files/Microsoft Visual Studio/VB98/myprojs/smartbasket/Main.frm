VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5520
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   220
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView BasketItems 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size (bytes)"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   5520
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Status 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0%"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label CurrentFile 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   5655
   End
   Begin VB.Label Items 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   490
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   525
      X2              =   0
      Y1              =   700
      Y2              =   700
   End
   Begin VB.Line Line3 
      X1              =   525
      X2              =   525
      Y1              =   480
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   525
      X2              =   0
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   525
      X2              =   525
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Image Basket 
      Height          =   435
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "Main.frx":0B24
      Top             =   0
      Width           =   525
   End
   Begin VB.Menu Pop 
      Caption         =   "&Pop"
      Visible         =   0   'False
      Begin VB.Menu PopItems 
         Caption         =   "&Compile file list"
      End
      Begin VB.Menu PopImp 
         Caption         =   "&Import"
      End
      Begin VB.Menu PopSep1 
         Caption         =   "-"
      End
      Begin VB.Menu PopTop 
         Caption         =   "&Always on top?"
         Checked         =   -1  'True
      End
      Begin VB.Menu PopMove 
         Caption         =   "&Moveable?"
         Checked         =   -1  'True
      End
      Begin VB.Menu PopSep2 
         Caption         =   "-"
      End
      Begin VB.Menu PopClose 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu PopRight 
      Caption         =   "&PopRight"
      Visible         =   0   'False
      Begin VB.Menu PopRightAll 
         Caption         =   "&Select all"
      End
      Begin VB.Menu PopRightDel 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ChangeBasket(mode As Integer)
    If mode = 0 Then
        Main.Width = 550
        Main.Height = 730
        BasketItems.Visible = False
    ElseIf mode = 1 Then
        Main.Width = 6030
        Main.Height = 5490
        BasketItems.Visible = True
    End If
End Sub


Private Sub CompileFile(FileName As String, files As ListView)
Dim k As Integer
Dim ifile As Long
Dim Dat As String, binary As String
Dim choice As Integer, starttime As Long, counter As Single
Dim num2inc As Long
Dim compiled As String
Dim stats As String

On Error GoTo FileError
    starttime = Timer
    For k = 1 To files.ListItems.Count
        ifile = FreeFile
        Open files.ListItems(k) For Binary As #ifile
            binary = "|||->" & files.ListItems(k) & "|||" & FileLen(files.ListItems(k)) & "<-|||"
            Dat = String(LOF(ifile), Chr(0))
            Get ifile, , Dat
        Close ifile
        binary = binary & Dat
        compiled = compiled & binary
        num2inc = Round(FileLen(files.ListItems(k)) / GetAllSize(BasketItems) * 100, 0)
        
        If Prog.Value + num2inc <= 100 Then
            Prog.Value = Prog.Value + Round(CInt(num2inc))
            Label3.Caption = ""
            Label3.Caption = Prog.Value & "%"
            Label3.Refresh
            CurrentFile.Caption = ""
            CurrentFile.Caption = "Processing " & GetFileFromPath(files.ListItems(k))
            CurrentFile.Refresh
        End If
    Next k
    ifile = FreeFile
    Open FileName For Output As #ifile
        Print #ifile, compiled
    Close ifile
    counter = Round(Timer, 2) - Round(starttime, 2)
    counter = Round(counter, 2)
    If counter = 0 Then counter = 0.5
    stats = ""
    stats = stats & "Time elapsed : " & counter & " secs" & vbCrLf
    stats = stats & "Files parsed : " & files.ListItems.Count & vbCrLf
    stats = stats & "Bytes read : " & GetAllSize(BasketItems) & " bytes" & " (" & Round(GetAllSize(BasketItems) / 1024, 0) & "K)" & vbCrLf
    stats = stats & "Compile speed : " & Round(((GetAllSize(BasketItems) / counter) / 1024), 2) & " k/ps" & vbCrLf
    choice = MsgBox(cdFile.FileTitle & " was written." & vbCrLf & vbCrLf & "Statistics : " & vbCrLf & "------------" & vbCrLf & stats & vbCrLf & vbCrLf & "Clear all basket data?", vbYesNo, "Compiling complete")
    
    Select Case choice
        Case 6
            BasketItems.ListItems.Clear
            Status.Caption = "0 files @ 0 bytes"
            Items.Caption = "0"
            If Main.Width > 1000 Then
                ChangeBasket (0)
            Else
                ChangeBasket (1)
            End If
    End Select
Prog.Value = 0
CurrentFile.Caption = ""
Exit Sub
FileError:
    MsgBox "[Error #" & Err.Number & " @ " & Err.Source & "] " & Err.Description, vbOKOnly, "Unexpected error"
    Prog.Value = 0
    CurrentFile.Caption = ""
End Sub

Private Sub ImportFile(FileName As String)
Dim k As Integer
Dim ifile As Long
Dim Dat As String, binary As String
Dim choice As Integer, starttime As Long, counter As Single
Dim num2inc As Long
Dim compiled As String
Dim stats As String
Dim pos As Long, pos2 As Long, pos3 As Long
Dim expfile As String
Dim newfile As Variant

On Error GoTo FileError
    starttime = Timer
    ifile = FreeFile
    Open FileName For Binary As #ifile
        Dat = String(LOF(ifile), Chr(0))
        Get ifile, , Dat
    Close ifile
    
        For pos = 1 To FileLen(FileName)
            If Mid(Dat, pos, 5) = "|||->" Then
                For pos2 = pos + 5 To FileLen(FileName)
                    expfile = expfile & Mid(Dat, pos2, 1)
                    If Mid(Dat, pos2, 5) = "<-|||" Then Exit For
                Next pos2
                If Mid(Dat, pos2, 5) = "<-|||" Then
                    expfile = Left(expfile, Len(expfile) - 1)
                    newfile = Split(expfile, "|||")
                    
                    If UBound(newfile) = "1" Then
                        BasketItems.ListItems.Add k + 1, , newfile(k), 1, 1
                        BasketItems.ListItems(k + 1).SubItems(1) = newfile(k + 1)
                        num2inc = Round(FileLen(BasketItems.ListItems(k + 1)) / pos2 * 100, 0)
                        If Prog.Value + num2inc <= 100 Then
                            Prog.Value = Prog.Value + Round(CInt(num2inc))
                            Label3.Caption = ""
                            Label3.Caption = Prog.Value & "%"
                            Label3.Refresh
                            CurrentFile.Caption = ""
                            CurrentFile.Caption = "Processing " & GetFileFromPath(BasketItems.ListItems(k + 1))
                            CurrentFile.Refresh
                        End If
                        expfile = ""
                        Erase newfile
                    End If
                End If
            End If
        Next pos
        Items.Caption = BasketItems.ListItems.Count
        Status.Caption = BasketItems.ListItems.Count & " files @ " & GetAllSize(BasketItems) & " bytes"
        
    counter = Round(Timer, 2) - Round(starttime, 2)
    counter = Round(counter, 1)
    If counter <= 0 Then counter = 0.1
    stats = ""
    stats = stats & "Time elapsed : " & counter & " secs" & vbCrLf
    stats = stats & "Files parsed : " & BasketItems.ListItems.Count & vbCrLf
    stats = stats & "Bytes read : " & GetAllSize(BasketItems) & " bytes" & " (" & Round(GetAllSize(BasketItems) / 1024, 0) & "K)" & vbCrLf
    stats = stats & "Compile speed : " & Round(((GetAllSize(BasketItems) / counter) / 1024), 2) & " k/ps" & vbCrLf
    If Prog.Value < 100 Then
        Prog.Value = 100
        Label3.Caption = "100%"
    End If
    choice = MsgBox(GetFileFromPath(FileName) & " was imported." & vbCrLf & vbCrLf & "Statistics : " & vbCrLf & "------------" & vbCrLf & stats, vbInformation, "Compiling complete")
    
Prog.Value = 0
CurrentFile.Caption = ""
Exit Sub
FileError:
    MsgBox "[Error #" & Err.Number & " @ " & Err.Source & "] " & Err.Description, vbOKOnly, "Unexpected error"
    Prog.Value = 0
    CurrentFile.Caption = ""
End Sub

Private Function GetAllSize(FileList As ListView) As Long
Dim size As Long
    GetAllSize = 0
    For k = 1 To FileList.ListItems.Count
        GetAllSize = GetAllSize + FileLen(FileList.ListItems(k))
    Next k
End Function

Private Sub Basket_DblClick()
    If Main.Width > 1000 Then
        ChangeBasket (0)
    Else
        ChangeBasket (1)
    End If
End Sub

Private Sub basket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 1 And PopMove.Checked = True Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&
    End If
End Sub

Private Sub basket_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If BasketItems.ListItems.Count > 0 Then
            PopItems.Enabled = True
        Else
            PopItems.Enabled = False
        End If
        Main.PopupMenu Pop
    End If
End Sub


Private Sub basket_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Integer, check As Integer
Dim choice As Integer
Dim size As Integer
    On Error GoTo droperror
    k = 1
    For Each file In Data.files
        If Dir(file) <> "" And Not vbDirectory Then
            BasketItems.ListItems.Add k, , file, 1, 1
            BasketItems.ListItems(k).SubItems(1) = FileLen(file)
        End If
    Next
        Items.Caption = BasketItems.ListItems.Count
        Status.Caption = BasketItems.ListItems.Count & " files @ " & GetAllSize(BasketItems) & " bytes"
Exit Sub
droperror:
Items.Caption = BasketItems.ListItems.Count
Status.Caption = BasketItems.ListItems.Count & " files @ " & GetAllSize(BasketItems) & " bytes"

End Sub



Private Sub BasketItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LCase(ColumnHeader) = "file name" Then
        BasketItems.SortKey = 0
        If BasketItems.SortOrder = lvwAscending Then
            BasketItems.SortOrder = lvwDescending
        Else
            BasketItems.SortOrder = lvwAscending
        End If
        BasketItems.Sorted = True
    End If
    
    If LCase(ColumnHeader) = "size (bytes)" Then
        BasketItems.SortKey = 1
        If BasketItems.SortOrder = lvwAscending Then
            BasketItems.SortOrder = lvwDescending
        Else
            BasketItems.SortOrder = lvwAscending
        End If
        BasketItems.Sorted = True
    End If
End Sub

Private Sub BasketItems_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Main.PopupMenu PopRight
    End If
End Sub


Private Sub BasketItems_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Integer, check As Integer
Dim choice As Integer
Dim size As Integer
    On Error GoTo droperror
    k = 1
    For Each file In Data.files
        If Dir(file) <> "" And Not vbDirectory Then
            BasketItems.ListItems.Add k, , file, 1, 1
            BasketItems.ListItems(k).SubItems(1) = FileLen(file)
        End If
    Next
        Items.Caption = BasketItems.ListItems.Count
        Status.Caption = BasketItems.ListItems.Count & " files @ " & GetAllSize(BasketItems) & " bytes"
Exit Sub
droperror:
Items.Caption = BasketItems.ListItems.Count
Status.Caption = BasketItems.ListItems.Count & " files @ " & GetAllSize(BasketItems) & " bytes"

End Sub


Private Sub BasketItems_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    Data.SetData , vbCFFiles
End Sub

Private Sub BasketItems_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    For Each ListItem In BasketItems.ListItems
        If ListItem.Selected Then
            Data.files.Add ListItem
        End If
    Next ListItem
    AllowedEffects = vbDropEffectCopy
    BasketItems_OLESetData Data, vbCFFiles
End Sub


Private Sub Form_Load()
    If LCase(getstring(HKEY_LOCAL_MACHINE, "Software\SmartBasket", "RunOnce")) <> "true" Then
        Associate "SmartBasket.Exec", ".sbe", "SmartBasket Executable", App.path & "\sbil.dll,1"
        Associate "SmartBasket.Zip", ".sbz", "SmartBasket ZipFile", App.path & "\sbil.dll,2"
        savestring HKEY_LOCAL_MACHINE, "Software\SmartBasket", "RunOnce", "true"
    End If
    
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Items.Caption = "0"
    Main.Height = 730
    Main.Width = 550
    BasketItems.Visible = False
    Status.Caption = "0 files @ 0 bytes"
    
    If Command <> "" Then
        Main.Show
        ChangeBasket (1)
        DoEvents
        ImportFile (Command)
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.path & "\images\compile_down.jpg")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.path & "\images\compile.jpg")
    PopItems_Click
End Sub


Private Sub Items_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And PopMove.Checked = True Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, _
        HTCAPTION, 0&
    End If
End Sub


Private Sub PopClose_Click()
    Unload Me
End Sub

Private Sub PopImp_Click()
Dim choice As Integer
    If BasketItems.ListItems.Count > 0 Then
        choice = MsgBox("Basket already contains items." & vbCrLf & "Clear list for import?", vbYesNoCancel, "Basket is not empty")
            If choice = vbYes Then
                Status.Caption = "0 files @ 0 bytes"
                Items.Caption = "0"
                BasketItems.ListItems.Clear
            ElseIf choice = vbCancel Then
                Exit Sub
            End If
    End If
                
    cdFile.Filter = "SmartBasket Zip (*.sbz)|*.sbz"
    cdFile.Action = 2
        
        If cdFile.FileName <> "" And Not cdFile.CancelError Then
            ImportFile cdFile.FileName
        End If
End Sub


Private Sub PopItems_Click()
On Error GoTo SaveError
Dim choice As Integer
    cdFile.FileName = ""
    cdFile.Filter = "SmartBasket Zip (*.sbz)|*.sbz"
    cdFile.Action = 2
        
        If cdFile.FileName = "" Then Exit Sub
        
        If Dir(cdFile.FileName) <> "" Then
            choice = MsgBox("File already exists." & vbCrLf & "Overwrite?", vbYesNo, "Overwrite?")
                If choice = vbNo Then
                    Exit Sub
                ElseIf choice = vbYes Then
                    CompileFile cdFile.FileName, Main.BasketItems
                End If
        Else
            CompileFile cdFile.FileName, Main.BasketItems
        End If
Exit Sub
SaveError:
    MsgBox "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub PopMove_Click()
    If PopMove.Checked Then
        PopMove.Checked = False
    Else
        PopMove.Checked = True
    End If
End Sub


Private Function GetFileFromPath(ByVal path As String) As String
Dim pos As Integer
    For pos = Len(path) To Mid(pos, 1, 1) Step -1
        If Mid(path, pos, 1) = "\" Then
            GetFileFromPath = Mid(path, pos + 1, Len(path))
            Exit For
        End If
    Next pos
End Function

Private Sub PopRightAll_Click()
    For Each ListItem In BasketItems.ListItems
        ListItem.Selected = True
    Next ListItem
End Sub

Private Sub PopRightDel_Click()
Dim k As Integer
Dim itemslist As Integer
    itemslist = BasketItems.ListItems.Count
    k = 1
    Do While k <= itemslist - 1
        If BasketItems.ListItems(k).Selected Then
            BasketItems.ListItems.Remove k
            Items.Caption = BasketItems.ListItems.Count
            k = k + 1
        End If
    Loop
End Sub

Private Sub PopTop_Click()
    If PopTop.Checked Then
        PopTop.Checked = False
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        PopTop.Checked = True
    End If
End Sub


