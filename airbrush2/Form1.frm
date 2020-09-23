VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airbrush using DIBits by Thomas Raben"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   20
      Left            =   4860
      Max             =   100
      TabIndex        =   10
      Top             =   3840
      Value           =   100
      Width           =   2115
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   20
      Left            =   1500
      Max             =   100
      Min             =   1
      TabIndex        =   9
      Top             =   3840
      Value           =   10
      Width           =   2235
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   420
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00AB846B&
      Height          =   375
      Index           =   6
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   1260
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Index           =   5
      Left            =   420
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   420
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   420
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   420
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   3795
      Left            =   840
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   0
      Top             =   0
      Width           =   6075
   End
   Begin VB.Label lblRender 
      Caption         =   "Rendertime:"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4560
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "For full performance, compile the project and run it :-)"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   4200
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "Pressure:"
      Height          =   255
      Left            =   4140
      TabIndex        =   12
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Radius:"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mouseDown As Boolean
Dim mColor As Long

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim radius As Long, tmp As String, beginTime As Long
    tmp = Me.HScroll1.Value
    
    'make sure we have an even radius...
    Select Case Right(tmp, 1)
        Case 1, 3, 5, 7, 9
            radius = CLng(tmp) + 1
        Case Else
            radius = CLng(tmp)
    End Select
    beginTime = Timer
    drawAirbrush Me.Picture1.hdc, CLng(X), CLng(Y), Me.HScroll1.Value, mColor, Me.HScroll2.Value
    Me.lblRender = "RenderTime: " & Abs(Timer - beginTime) & " seconds."
    mouseDown = True
    
End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim beginTime As Long
    
    If mouseDown Then
        beginTime = Timer
        drawAirbrush Me.Picture1.hdc, CLng(X), CLng(Y), Me.HScroll1.Value, mColor, Me.HScroll2.Value
        Me.lblRender = "RenderTime: " & Abs(Timer - beginTime) & " seconds."
    End If
    
End Sub

Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown = False
    
End Sub

Private Sub Picture2_Click(Index As Integer)
    mColor = Me.Picture2(Index).BackColor
    
End Sub
