VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Grafiken aufrollen"
   ClientHeight    =   8685
   ClientLeft      =   1290
   ClientTop       =   1605
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   10845
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture2 
      Height          =   8655
      Left            =   0
      ScaleHeight     =   573
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   717
      TabIndex        =   1
      Top             =   0
      Width           =   10815
      Begin MSComDlg.CommonDialog CD 
         Left            =   4440
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   8160
         TabIndex        =   2
         Top             =   5880
         Width           =   2535
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            Max             =   200
            Min             =   1
            TabIndex        =   9
            Top             =   2160
            Value           =   1
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   120
            Max             =   200
            Min             =   1
            TabIndex        =   7
            Top             =   1440
            Value           =   50
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Start"
            Height          =   425
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Exit"
            Height          =   425
            Left            =   1200
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Load file"
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Effect"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Size:"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   1200
            Width           =   2055
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Copyright ©2003 by  by CC
'//  Email: cyber_chris235@gmx.net
'//   All Rights reserved

Option Explicit



Private Sub Command1_Click()
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double, TubeWidth
    TubeWidth = HScroll1.Value
    Picture2.Cls
    RDS = 6.283185307 / (TubeWidth * 2) * HScroll2.Value
    For Offset = 0 To Picture1.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then Picture2.PaintPicture Picture1.Picture, Offset - TubeWidth, 0, 1, Picture1.ScaleHeight, Offset - TubeWidth, 0, 1, Picture1.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Picture1.ScaleWidth Then
                Picture2.PaintPicture Picture1.Picture, Offset + XTube - TubeWidth, 0, 1, Picture1.ScaleHeight, Offset + XPicture, 0, 1, Picture1.ScaleHeight
            Else
                Picture2.PaintPicture Picture1.Picture, Offset + XTube - TubeWidth, 0, 1, Picture1.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Picture1.ScaleHeight
            End If
        Next XTube
    Next Offset
End Sub

Private Function ACos(X As Double)
    X = X - 1
    If X < 1 And X > -1 Then
        ACos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    Else
        ACos = 0
    End If
End Function

Private Sub Command3_Click()
Dim strfilename
With CD
    .DialogTitle = "Load Picture"
    .Filter = "Bitmap" & "(*.bmp)|*.bmp|Jpeg Files (*.jpg)|*.jpg"
    .FilterIndex = 1
    .ShowOpen
    .CancelError = False
    strfilename = .FileName
    Picture1.Picture = LoadPicture(strfilename)
End With
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    Picture2.Width = Form1.Width
    Picture2.Height = Form1.Height
End Sub

Private Sub Command2_Click()
  Unload Me
  End
End Sub

Private Sub Form_Resize()
    Picture2.Width = Form1.Width
    Picture2.Height = Form1.Height
End Sub

