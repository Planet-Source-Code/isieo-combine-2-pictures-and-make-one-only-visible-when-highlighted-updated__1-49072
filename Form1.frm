VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PICTURE HIDER!!"
   ClientHeight    =   3915
   ClientLeft      =   150
   ClientTop       =   885
   ClientWidth     =   5685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CBFix 
      Caption         =   "Fit to box"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   750
      Width           =   2175
   End
   Begin VB.TextBox Txthtml 
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "&Stop"
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar PBar 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3015
      Left            =   3000
      ScaleHeight     =   2955
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   3240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton CmdLoad2 
      Caption         =   "&Load The invisible Picture "
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton CmdLoad1 
      Caption         =   "&Load The Visible Picture "
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   2880
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton CmdMix 
      Caption         =   "&Embed In!!!"
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BoolStop As Boolean
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Private Sub CBFix_Click()
On Error GoTo err:
If CBFix.Value = 1 Then ' when the check box is check then
    Picture1.Height = 2895 'set initial size
    Picture1.Width = 2775 'set initial size
    'Calculation for resizing the picturebox and picture
    If Picture1.Picture.Width <= Picture1.Picture.Height Then
        Picture1.Width = (Picture1.Height / Picture1.Picture.Height) * Picture1.Picture.Width
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0   'Resize the picture acording to scale
    Else
        Picture1.Height = (Picture1.Width / Picture1.Picture.Width) * Picture1.Picture.Height
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0      'Resize the picture acording to scale
    End If
End If
    Form_Resize
Exit Sub
err:
CBFix.Value = 1
End Sub

Private Sub CmdMix_Click()
'On Error GoTo Oho
Dim x As Long
Dim y As Long
Dim Fileloc As String
CBFix.Enabled = False ' gui
CmdMix.Enabled = False 'gui
CmdLoad1.Enabled = False 'gui
CmdLoad2.Enabled = False 'gui
CmdStop.Visible = True 'gui
BoolStop = False 'set the boolean false
cdg1.FileName = Left(cdg1.FileName, Len(cdg1.FileName) - 4) 'remove extention
cdg1.ShowSave 'shows where to save the new picture
Fileloc = cdg1.FileName
Picture3.Cls  'clear the picture box
Picture3.Height = Picture1.Height 'resize th picture box
Picture3.Width = Picture1.Width 'resize
    For y = 0 To Picture1.Height 'loop! until y is same as the height of the picture
        DoEvents
        If BoolStop = True Then 'check if user press stop button
            Select Case MsgBox("Do you want to save it so far?", vbYesNoCancel)
                Case vbYes 'user wants to save
                    Exit For
                Case vbNo 'user don't want to save
                    GoTo Oho:
                Case vbCancel 'User wants it to continue
                    BoolStop = False
                End Select
        End If
        For x = 0 To Picture1.Width  'another loop but now its x.
                ' when the x loop finishes, x will become 0 again and y will add 1 which means,
                  'we are filling in the pixels line by line
'***###IMPOTRANT!!###*** ALL THIS WILL ONLY WORK IF THE PICTURE BOX IS SET AUTOREDRAW = TRUE
            If x Mod 2 <> 0 Then ' this is to alternate the mix of the pictures in the x axis
                If y Mod 2 <> 0 Then ' this is to alternate the mix of the pictures in the y axis
                SetPixelV Picture3.hdc, x, y, GetPixel(Picture1.hdc, x, y) ' set the pixel with the selected colours
                Else
                SetPixelV Picture3.hdc, x, y, GetPixel(Picture2.hdc, x, y) ' set the pixel with the selected colours
                End If
            Else
                If y Mod 2 = 0 Then ' this is to alternate the mix of the pictures in the y axis
                SetPixelV Picture3.hdc, x, y, GetPixel(Picture1.hdc, x, y) ' set the pixel with the selected colours
                Else
                SetPixelV Picture3.hdc, x, y, GetPixel(Picture2.hdc, x, y) ' set the pixel with the selected colours
                End If
            End If
        Next x
        Picture3.Picture = Picture3.Image 'convert the image into the picture
        Image1.Picture = Picture3.Picture ' show the preview
        Me.Caption = "PICTURE HIDER!! " & PBar.Value & "% Done"   ' Display the percentage
        If PBar.Value > 97 Then Exit For
        PBar.Value = Round((y / (Picture1.Height)) * 1000, 0)   ' working for progressbar
          Next y
PBar.Value = 0 ' set it to empty
Me.Caption = "PICTURE HIDER!!" ' Change the form caption
Picture3.Picture = Picture3.Image ' to reasure that the image is converted to the picture
SavePicture Picture3.Image, Fileloc & ".bmp" 'save the picture in to a file
'Create HTML
If MsgBox("Do you want to create a HTML file to view it?", vbYesNo, "HTML??") = vbYes Then
    Txthtml.Text = "<html>" & vbCrLf & "<body bgcolor=" & Chr(34) & "#000000" & Chr(34) & "text=" & Chr(34) & "#00FF00" & Chr(34) & ">" & vbCrLf _
               & "<center><img src=" & Chr(34) & Fileloc & ".bmp" & Chr(34) & "> <p> Highlight the picture!!!" & vbCrLf & " </center></body></html>"
    Open Fileloc & ".html" For Binary Access Write As #1
    Put #1, , Txthtml.Text
    Close #1
End If
Oho:
CBFix.Enabled = True
CmdLoad1.Enabled = True
CmdMix.Enabled = True
CmdStop.Visible = False
End Sub

Private Sub CmdLoad1_Click()
On Error GoTo Oho
cdg1.ShowOpen ' show the open dilog box
If CBFix.Value = 1 Then
    Picture1.Height = 2895
    Picture1.Width = 2775
    '**##IMPORTANT!!##** THIS WILL ONLY WORK WITH AUTOREDRAW = TRUE
    Picture1.AutoSize = False
    Picture1.Picture = LoadPicture(cdg1.FileName) 'loads the picture
        'Calculation for resizing the picturebox and picture
        If Picture1.Picture.Width <= Picture1.Picture.Height Then
        Picture1.Width = (Picture1.Height / Picture1.Picture.Height) * Picture1.Picture.Width
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0   'Resize the picture acording to scale
    Else
        Picture1.Height = (Picture1.Width / Picture1.Picture.Width) * Picture1.Picture.Height
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height, 0, 0      'Resize the picture acording to scale
    End If
    Picture1.Picture = Picture1.Image 'loads the image into the picture
    Picture1.AutoSize = True
Else
    Picture1.Picture = LoadPicture(cdg1.FileName) 'loads the picture
End If
Form_Resize 'resize the form
CmdLoad2.Enabled = True 'enable the command button
Exit Sub
Oho:
Beep
MsgBox Error, , "Error"
End Sub

Private Sub CmdLoad2_Click()
On Error GoTo Oho
cdg1.ShowOpen ' show the open dilog box
Picture2.Picture = LoadPicture(cdg1.FileName) 'loads the picture
'**##IMPORTANT!!##** THIS WILL ONLY WORK WITH AUTOREDRAW = TRUE
Picture2.PaintPicture Picture2.Picture, 0, 0, Picture2.Width, Picture2.Height, 0, 0 'Resize the picture
Picture2.Picture = Picture2.Image 'loads the image into the picture
Form_Resize 'resize the form
CmdMix.Enabled = True 'enable the command button
Exit Sub
Oho:
Beep
MsgBox Error, , "Error"
End Sub
Private Sub CmdStop_Click()
BoolStop = True
End Sub

Private Sub Form_Resize()
' calculation to resize the form
Picture2.Width = Picture1.Width
Picture2.Height = Picture1.Height
Picture2.Left = Picture1.Left + Picture1.Width + 50
If Picture1.Width < 2775 Then
    Me.Width = Picture1.Width + Picture2.Width + 2775
Else
    Me.Width = Picture1.Width + Picture2.Width + 200
End If
If Picture1.Height < 2895 Then
    Me.Height = Picture1.Height + Picture1.Top + 2895
    Else
    Me.Height = Picture1.Height + Picture1.Top + 900
End If
PBar.Width = Me.Width - (CmdMix.Width + CmdLoad1.Width + Image1.Width) - 120
End Sub

Private Sub MnuExit_Click(Index As Integer)
Unload Me
End
End Sub
