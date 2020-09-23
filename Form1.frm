VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Ini Explorer"
   ClientHeight    =   4560
   ClientLeft      =   2130
   ClientTop       =   2070
   ClientWidth     =   5805
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5805
   Begin VB.CommandButton Command7 
      Caption         =   "Save As"
      Height          =   555
      Left            =   4995
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   690
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete Node"
      Height          =   555
      Left            =   4215
      Picture         =   "Form1.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3855
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "New"
      Height          =   555
      Left            =   4200
      Picture         =   "Form1.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "New Node"
      Height          =   555
      Left            =   4215
      Picture         =   "Form1.frx":0B78
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3255
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Expand All"
      Height          =   555
      Left            =   4215
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":0C7A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2685
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4245
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   555
      Left            =   4170
      Picture         =   "Form1.frx":0D64
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   555
      Left            =   4200
      Picture         =   "Form1.frx":1296
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   4230
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   7461
      _Version        =   327682
      Indentation     =   793
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selected Item Key = "
      Height          =   195
      Left            =   165
      TabIndex        =   5
      Top             =   4350
      Width           =   1470
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4245
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":17C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2116
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public File As String, DaBa As String, Nod As Node, RNod As Node, CurNew As String, NewNod As Integer

Private Sub Command1_Click()
CD.Filter = "INI Files|*.ini"
CD.ShowOpen
If CD.FileName = "" Then Exit Sub
File = CD.FileName
LoadINI
End Sub

Sub LoadINI()
DaBa = ""
Dim X, Y, Z, GenKey As Integer, A, B, CurInfo As String, CurData As String, CurDir As String, CurDirPos
TreeView1.Nodes.Clear
Set RNod = TreeView1.Nodes.Add(, , File, File, 1)
DaBa = String(FileLen(File), " ")
Open File For Binary As #1
Get #1, 1, DaBa
Close #1

For Y = 1 To Len(DaBa)
    If Mid(DaBa, Y, 1) = "[" Then
        For X = Y To Len(DaBa)
        
        If Mid(DaBa, X, 1) = "]" Then
        CurDirPos = Y + 1
        CurDir = Mid(DaBa, Y + 1, (X - Y) - 1)
        On Error Resume Next
        Set Nod = TreeView1.Nodes.Add(File, tvwChild, "Root" & CurDir, CurDir, 2)
 
           For Z = X + 1 To Len(DaBa)
           If Mid(DaBa, Z, 1) = "[" Then Exit For
           If Mid(DaBa, Z, 1) = "=" Then
               For A = Z To 1 Step -1
               If Mid(DaBa, A, 1) = "]" Then Exit For
                If Mid(DaBa, A, 1) = Chr(13) Then
                CurInfo = Mid(DaBa, A + 2, Z - A - 2)
                Set Nod = TreeView1.Nodes.Add("Root" & CurDir, tvwChild, "Info" & CurInfo & CurDir, CurInfo, 3)
                Exit For
                End If
               Next A
               For A = Z To Len(DaBa)
                If Mid(DaBa, A, 1) = "[" Then Exit For
                If Mid(DaBa, A, 1) = Chr(13) Or A = Len(DaBa) Then
                If A = Len(DaBa) Then Let A = A + 1
                CurData = Mid(DaBa, Z + 1, A - (Z + 1))
                Set Nod = TreeView1.Nodes.Add("Info" & CurInfo & CurDir, tvwChild, , CurData, 4)
                Exit For
                End If
                Next A
            
             End If
             Next Z
        Exit For
        End If
        Next X
    End If

Next Y
TreeView1.Nodes.Item(1).Expanded = True
End Sub

Private Sub Command2_Click()
Dim Reply
Reply = MsgBox("This will erase the old file and build up the new one from the tree, not including blank headings....", vbOKCancel)
If Reply = vbCancel Then Exit Sub
On Error Resume Next
Kill (File)
Open File For Output As #1
Close #1
Dim Node As Node
For Each Node In TreeView1.Nodes
If Mid(Node.Key, 1, 4) = "Info" Then
Label1.Caption = "Now writing " & Node.Text
WriteINI File, Node.Parent.Text, Node.Text, Node.Child.Text
End If
Next
End Sub

Private Sub Command3_Click()
For I = 1 To TreeView1.Nodes.Count
TreeView1.Nodes.Item(I).Expanded = True
Next
End Sub

Private Sub Command4_Click()
Dim Nod As Node, Img As Integer
NewNod = NewNod + 1
If TreeView1.SelectedItem.Key = File Then
Set Nod = TreeView1.Nodes.Add(File, tvwChild, "RootNew Node" & NewNod, "New Node" & NewNod, 2)
TreeView1.Nodes.Item(1).Expanded = True
Exit Sub
End If
If TreeView1.SelectedItem.Key = "" Then Exit Sub
If Mid(TreeView1.SelectedItem.Key, 1, 4) = "Root" Then
Set Nod = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "Info" & "New Node" & NewNod & TreeView1.SelectedItem.Text, "New Node" & NewNod, 3)
Set Nod = TreeView1.Nodes.Add("Info" & "New Node" & NewNod & TreeView1.SelectedItem.Text, tvwChild, "", "Nothing", 4)
End If
End Sub

Private Sub Command5_Click()
CD.FileName = ""
CD.Filter = "Ini Files|*.ini"
CD.ShowSave
If CD.FileName = "" Then Exit Sub
File = CD.FileName
Open File For Output As #1
Close #1
TreeView1.Nodes.Clear
Set Nod = TreeView1.Nodes.Add(, , File, File, 1)
End Sub

Private Sub Command6_Click()
If TreeView1.SelectedItem.Key = "" Then Beep: Exit Sub
TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
End Sub

Private Sub Command7_Click()
CD.FileName = ""
CD.ShowSave
If CD.FileName = "" Then Exit Sub
File = CD.FileName
Command2_Click
End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
If TreeView1.SelectedItem.Key = File Then Cancel = True: Exit Sub
If TreeView1.SelectedItem.Key = "" Then Exit Sub
If TreeView1.SelectedItem.Parent.Key = File Then Let TreeView1.SelectedItem.Key = "Root" & NewString: Exit Sub
If Mid(TreeView1.SelectedItem.Parent.Key, 1, 4) = "Root" Then Let TreeView1.SelectedItem.Key = "Info" & NewString & TreeView1.SelectedItem.Parent.Text: Exit Sub
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Let Label1.Caption = "Key = " & Node.Key
End Sub
Private Function WriteINI(FileName1 As String, Heading1 As String, Variable1 As String, Value1 As String)
WritePrivateProfileString Heading1, Variable1, Value1, FileName1
For I = 0 To 2000
DoEvents
Next I
End Function
