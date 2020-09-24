VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucComboBoxEx ucComboBoxEx2 
      Height          =   2370
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4180
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sty             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin Project1.ucComboBoxEx ucComboBoxEx1 
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1440
      Width           =   8655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   571
      TabIndex        =   0
      Top             =   2400
      Width           =   8595
   End
   Begin VB.Label Label2 
      Caption         =   "MS Combo:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "Unicode Simple ucComboBoxEx using SimSum Font:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Unicode ucComboBoxEx using Tahoma Font:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a PictureBox that is used to display Unicode via DrawTextW"
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2700
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgFlags 
      Height          =   255
      Left            =   720
      Picture         =   "Form1.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   5865
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type RECT
    Left                 As Long
    Top                  As Long
    Right                As Long
    Bottom               As Long
End Type
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function DrawTextW Lib "user32" (ByVal hDc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private m_cFlg16        As New cImageList 'Country flags
Private m_Rect          As RECT

Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long

Private Const MB_ICONINFORMATION As Long = &H40&
Private Const MB_TASKMODAL As Long = &H2000&

Private Sub Command1_Click()
   
   '#If Unicode Then
         MessageBoxW Me.hWnd, StrPtr(LoadResString(101 + 2) & "|" & LoadResString(121 + 2)), StrPtr("The string you are trying to find is:"), MB_ICONINFORMATION Or MB_TASKMODAL
         MsgBox "Index is   " & ucComboBoxEx1.FindItem_Unicode(LoadResString(101 + 2) & "|" & LoadResString(121 + 2), False)

   '#Else
      'If ucComboBoxEx1.FindItem(LoadResString(101 + 2) & "|" & LoadResString(121 + 2), False) > -1 Then
         'MsgBox "Index is   " & ucComboBoxEx1.FindItem(LoadResString(101 + 2) & "|" & LoadResString(121 + 2), False)
      'Else
         'MsgBox "Failed to find"
      'End If
   '#End If
   
End Sub

Private Sub Form_Initialize()
   'Put here,otherwise,XP theme MANIFEST may get problem
   InitCommonControls
End Sub

Private Sub Form_Load()

Dim i                As Long

    m_cFlg16.CreateFromHandle ILC_COLOR8, 23, 16, _
                              imgFlags.Picture.Handle, vbMagenta
    SetRect m_Rect, 8 + m_cFlg16.IconSizeX, 4, Picture1.Width, Picture1.Height
    With ucComboBoxEx1
        'Call .AddBitmap(LoadResPicture("TOOLBAR", vbResBitmap), vbMagenta)
        .ImageList = m_cFlg16.himl
        '.Font.Name = "Tahoma"
        '.Font.Size = 12
        For i = 0 To 16        'add 17 test items
            .AddItem LoadResString(101 + i) & "|" & LoadResString(121 + i), , i, i
            Combo1.AddItem LoadResString(101 + i) & "|" & LoadResString(121 + i)
        Next i
        .ListIndex = 1
        Combo1.ListIndex = 1
    End With
    
      With ucComboBoxEx2
        .ImageList = m_cFlg16.himl
        'Call .AddBitmap(LoadResPicture("TOOLBAR", vbResBitmap), vbMagenta)
        For i = 0 To 16        'add 17 test items
            .AddItem LoadResString(101 + i) & "|" & LoadResString(121 + i), , i, i
        Next i
        .ListIndex = 1
    End With
End Sub

Private Sub ucComboBoxEx1_ListIndexChange()

Dim s    As String
Dim lPtr As Long
    
    s = ucComboBoxEx1.Text
    lPtr = StrPtr(s)
    Picture1.Cls
    If lPtr Then
        m_cFlg16.Draw ucComboBoxEx1.ListIndex, Picture1.hDc, 4, 2
        DrawTextW Picture1.hDc, lPtr, Len(s), m_Rect, 0
    End If

End Sub

Private Sub ucComboBoxEx2_Click()
    Dim s    As String
Dim lPtr As Long
    s = ucComboBoxEx2.Text
    lPtr = StrPtr(s)
    Picture1.Cls
    If lPtr Then
        m_cFlg16.Draw ucComboBoxEx2.ListIndex, Picture1.hDc, 4, 2
        DrawTextW Picture1.hDc, lPtr, Len(s), m_Rect, 0
    End If
End Sub



