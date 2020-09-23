VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   Caption         =   "Crowd Sim"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   716
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSHUFFLE 
      Caption         =   "shuffle"
      Height          =   375
      Left            =   13920
      TabIndex        =   7
      Top             =   6360
      Width           =   1095
   End
   Begin VB.HScrollBar sTS 
      Height          =   255
      Left            =   13920
      Max             =   10
      Min             =   2
      TabIndex        =   6
      Top             =   5160
      Value           =   10
      Width           =   1095
   End
   Begin VB.HScrollBar sTL 
      Height          =   255
      Left            =   13920
      Max             =   50
      TabIndex        =   5
      Top             =   4800
      Value           =   5
      Width           =   1095
   End
   Begin VB.HScrollBar sZOOM 
      Height          =   255
      Left            =   13920
      Max             =   100
      Min             =   25
      TabIndex        =   4
      Top             =   3600
      Value           =   65
      Width           =   1095
   End
   Begin VB.ComboBox cmbTargetMode 
      Height          =   315
      Left            =   13920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtNH 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   2
      Text            =   "200"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   1335
      Left            =   13920
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   120
      ScaleHeight     =   663
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   903
      TabIndex        =   0
      Top             =   240
      Width           =   13575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trail"
      Height          =   255
      Left            =   13920
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N Agents"
      Height          =   255
      Left            =   13920
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lZOOM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ZOOM"
      Height          =   255
      Left            =   13920
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'Crowd Simulation
'Virtual Walkers
'realistic walking in virtual world


Option Explicit
Public T           As Long

Private Sub cmbTargetMode_Click()
    Dim I          As Long
    Dim AA         As Single

    For I = 1 To NH
        ReTarget I, False
        ReColor I
    Next

End Sub

Private Sub cmdSHUFFLE_Click()
Dim I As Long
For I = 1 To NH
H(I).X = Rnd * MaxX
H(I).Y = Rnd * MaxY
Next I
End Sub

Private Sub Command_Click()
    InitH

    Do
        If T Mod 4 = 0 Then DRAW
        MoveMent
        DoEvents
        T = T + 1
        If T Mod TrailStep = 0 Then ADDTrail
    Loop While True

End Sub

Private Sub Form_Load()

                    ZOOM = sZOOM / 100
                    
    rZOOM = r * ZOOM
   
    Randomize Timer

    MaxX = PIC.Width / ZOOM
    MaxY = PIC.Height / ZOOM


    cmbTargetMode.AddItem "Circle"
    cmbTargetMode.AddItem "Front"
    cmbTargetMode.AddItem "Perp"
    cmbTargetMode.AddItem "Cross"
     cmbTargetMode.AddItem "Side-Side"

    cmbTargetMode.ListIndex = 0

    TrailLen = sTL
    TrailStep = sTS
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sTL_Change()
   TrailLen = sTL
    
    
End Sub

Private Sub sTS_Change()

    TrailStep = sTS
    
End Sub

Private Sub sZOOM_Change()
    Dim I          As Long

    ZOOM = sZOOM / 100
    rZOOM = r * ZOOM
    MaxX = PIC.Width / ZOOM
    MaxY = PIC.Height / ZOOM
    For I = 1 To NH
        If cmbTargetMode.ListIndex <> 0 Then ReTarget I, False
    Next

End Sub

Private Sub txtNH_KeyPress(KeyAscii As Integer)
    Dim nnH        As Long
    Dim I          As Long
    If KeyAscii = 13 Then
        nnH = txtNH
        If nnH < NH Then
            NH = nnH
            ReDim Preserve H(NH)
        Else
            ReDim Preserve H(nnH)
            For I = 1 To NH: ReTarget I, False: ReColor I: Next
            For I = NH + 1 To nnH
                With H(I)
                    .Vstd = 1 + Rnd * 0.2

                    ReTarget I, True: ReColor I
                End With
            Next
            NH = nnH
        End If



    End If

End Sub


