VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Turbulance File Browser"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "desktop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox bar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00F97346&
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   0
      ScaleHeight     =   7470
      ScaleWidth      =   1965
      TabIndex        =   4
      Top             =   0
      Width           =   1995
      Begin VB.PictureBox prog 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1080
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":0894
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":0CE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":1138
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":188A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":1CDC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   360
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":212E
               Key             =   "openfolder"
               Object.Tag             =   "openfolder"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":2580
               Key             =   "internet"
               Object.Tag             =   "internet"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":29D2
               Key             =   "closedfolder"
               Object.Tag             =   "closedfolder"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":2E24
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":3576
               Key             =   "txt"
               Object.Tag             =   "txt"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":39C8
               Key             =   "bmp"
               Object.Tag             =   "bmp"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":3E1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "desktop.frx":426C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label infolabel 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image ico 
         Height          =   495
         Left            =   0
         Top             =   120
         Width           =   495
      End
      Begin VB.Label file 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick A File"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   750
      End
      Begin VB.Label folder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C:\"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   195
      End
      Begin VB.Image preview 
         Height          =   1455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label infolabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Pre-View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label infolabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView filelist 
      Height          =   6855
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   12091
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'stuff to grab the exe icons
Private Type PicBmp
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect _
Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" _
Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal _
nIconIndex As Long, phiconLarge As Long, phiconSmall As _
Long, ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal _
hicon As Long) As Long
Public Function GetIconFromFile(FileName As String, _
IconIndex As Long, UseLargeIcon As Boolean) As Picture

'Parameters:
'FileName - File (EXE or DLL) containing icons
'IconIndex - Index of icon to extract, starting with 0
'UseLargeIcon-True for a large icon, False for a small icon
'Returns: Picture object, containing icon

Dim hlargeicon As Long
Dim hsmallicon As Long
Dim selhandle As Long

' IPicture requires a reference to "Standard OLE Types."
Dim pic As PicBmp
Dim IPic As IPicture
Dim IID_IDispatch As GUID

If ExtractIconEx(FileName, IconIndex, hlargeicon, _
hsmallicon, 1) > 0 Then

selhandle = hlargeicon


' Fill in with IDispatch Interface ID.
With IID_IDispatch
.Data1 = &H20400
.Data4(0) = &HC0
.Data4(7) = &H46
End With
' Fill Pic with necessary parts.
With pic
.Size = Len(pic) ' Length of structure.
.tType = vbPicTypeIcon ' Type of Picture (bitmap).
.hBmp = selhandle ' Handle to bitmap.
End With

' Create Picture object.
Call OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

' Return the new Picture object.
Set GetIconFromFile = IPic

DestroyIcon hsmallicon
DestroyIcon hlargeicon

End If
End Function
Private Function fillasdrive()
Dim drive
filelist.ListItems.Clear
For x = 0 To Drive1.ListCount - 1
drive = UCase(Drive1.List(x)) + "\"
drive = "<" + Mid$(drive, 1, (Len(drive))) + ">"
filelist.ListItems.Add , , drive, 8
Next x
End Function

Private Function FillFileList()
'On Error Resume Next
'for folders

Dim x, i, ex
filelist.ListItems.Add , , "<..Drives..>", 8
If Right$(Dir1.Path, 2) <> ":\" Then filelist.ListItems.Add , , "<..>", 7

 For x = 0 To Dir1.ListCount - 1
If Right$(Dir1.Path, 2) <> ":\" Then newdir = "<" + Right(Dir1.List(x), Len(Dir1.List(x)) - Len(Dir1.Path) - 1) + ">" Else newdir = "<" + Right(Dir1.List(x), Len(Dir1.List(x)) - Len(Dir1.Path)) + ">"

 filelist.ListItems.Add , , newdir, 3, 3
 Next x
'for files


For i = 1 To File1.ListCount - 1
x = File1.List(i)
ex = UCase$(Mid$(x, Len(x) - 2, 4))
ex2 = UCase$(Mid$(x, Len(x) - 3, 5))
    
    
    If ex = "EXE" Then
    Dim z As String
z = Dir1 + "\" + File1.List(i)
    Set prog.Picture = GetIconFromFile(z, 0, True)

    If prog.Picture = 0 Then GoTo error
    ImageList1.ListImages.Add , , prog.Image
    ImageList2.ListImages.Add , , prog.Image
    a = ImageList1.ListImages.Count
    b = ImageList2.ListImages.Count
        filelist.ListItems.Add , , x, a, b
error:
d = 1
    End If
    
    If ex = "TXT" Then filelist.ListItems.Add , , x, 5, 5: d = 1
    If ex = "DOC" Then filelist.ListItems.Add , , x, 5, 5: d = 1
    If ex = "RTF" Then filelist.ListItems.Add , , x, 5, 5: d = 1
    If ex = "BMP" Then filelist.ListItems.Add , , x, 6, 6: d = 1
    If ex = "JPG" Then filelist.ListItems.Add , , x, 6, 6: d = 1
    If ex = "GIF" Then filelist.ListItems.Add , , x, 6, 6: d = 1
    If ex = "HTM" Or ex2 = "HTML" Then filelist.ListItems.Add , , x, 2, 2: d = 1
    If d <> 1 Then filelist.ListItems.Add , , x, 4, 4



Next i
End Function

Private Sub Command1_Click()
filelist.ListItems.Clear
FillFileList
End Sub

Private Sub Dir1_Change()
File1 = Dir1
If Right$(Dir1, 1) = "\" Then folder = Dir1 Else folder = Dir1 + "\"
filelist.ListItems.Clear
FillFileList
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1

End Sub

Private Sub File1_Click()
file = File1

End Sub

Private Sub filelist_DblClick()
On Error Resume Next
'if the drives button is pressed
If filelist.SelectedItem.Text = "<..Drives..>" Then fillasdrive: Exit Sub
'if back is pressed
'If filelist.SelectedItem.Text = "<..>" Then File1.Path = "..": Exit Sub


'If its A Folder then
If Left$(filelist.SelectedItem.Text, 1) = "<" Then
    'take the <> off
    folder = Mid$(filelist.SelectedItem.Text, 2, (Len(filelist.SelectedItem.Text) - 2))
    Dir1.Path = folder
    'goto that folder

End If

'find the filename
x = filelist.SelectedItem.Text
'find the extention
ex = UCase$(Mid$(x, Len(x) - 2, 4))

'if the files not a folder and is an exe file then
If Left$(filelist.SelectedItem.Text, 1) <> "<" Then
'if its an EXE File then run it
If ex = "EXE" Then Shell folder + file, vbNormalFocus

End If
End Sub


Private Sub filelist_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim x As String
On Error Resume Next
x = filelist.SelectedItem.Text
If Left$(x, 1) = "<" Then Exit Sub

file = x

'preview
x = filelist.SelectedItem.Text
ex = UCase$(Mid$(x, Len(x) - 2, 4))
If ex = "BMP" Or "GIF" Or "JPG" Then
infolabel(6).Visible = True
preview.Visible = True
preview.Picture = LoadPicture(folder + file)
Else
infolabel(6).Visible = False
preview.Visible = False
End If

' icon

    If ex <> "EXE" Then ico.Picture = ImageList1.ListImages(filelist.SelectedItem.Icon).Picture
    If ex = "EXE" Then ico.Picture = GetIconFromFile(folder + file, 0, True)
End Sub

Private Sub Form_Load()
fillasdrive


End Sub

Private Sub Label2_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Form_Resize()
filelist.Width = Me.Width - 2150
filelist.Height = Me.Height - 350
End Sub
