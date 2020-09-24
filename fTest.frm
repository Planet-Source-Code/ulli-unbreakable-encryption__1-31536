VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fTest 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Variable Encryption Test"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   ForeColor       =   &H8000000F&
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8340
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ImageList imlNodeTypes 
      Left            =   3705
      Top             =   5415
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":0CCE
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":0EDE
            Key             =   "Leaf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest.frx":10EE
            Key             =   "Node"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwTree 
      Height          =   2325
      Left            =   900
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4725
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   4101
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlNodeTypes"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txHex 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3675
      Width           =   3555
   End
   Begin VB.ListBox lstCodes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   4590
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   585
      Width           =   3510
   End
   Begin VB.TextBox txScrambled 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2625
      Width           =   3555
   End
   Begin VB.TextBox txClear 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   930
      Width           =   3555
   End
   Begin VB.CommandButton btDecrypt 
      Caption         =   "&Decrypt"
      Height          =   495
      Left            =   3390
      TabIndex        =   5
      Top             =   1995
      Width           =   1065
   End
   Begin VB.TextBox txKey 
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Top             =   210
      Width           =   3510
   End
   Begin VB.CommandButton btEncrypt 
      Caption         =   "&Encrypt"
      Height          =   495
      Left            =   900
      TabIndex        =   4
      Top             =   1995
      Width           =   1065
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      Caption         =   "Tree"
      Height          =   195
      Index           =   7
      Left            =   465
      TabIndex        =   17
      Top             =   4770
      Width           =   330
   End
   Begin VB.Label lbCount 
      Alignment       =   2  'Zentriert
      Height          =   195
      Left            =   2265
      TabIndex        =   15
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      Caption         =   "Encypted Text (hex)"
      Height          =   420
      Index           =   6
      Left            =   90
      TabIndex        =   14
      Top             =   3675
      Width           =   735
   End
   Begin VB.Label lb 
      Caption         =   "Generated Codes"
      Height          =   195
      Index           =   5
      Left            =   4590
      TabIndex        =   12
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label lbMinMax 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   900
      TabIndex        =   10
      Top             =   525
      Width           =   45
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ò"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   1950
      TabIndex        =   8
      Top             =   2100
      Width           =   210
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ñ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   3150
      TabIndex        =   9
      Top             =   2085
      Width           =   255
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      Caption         =   "Encypted Text"
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2670
      Width           =   690
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      Caption         =   "&Clear Text"
      Height          =   390
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   945
      Width           =   375
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      Caption         =   "&Key"
      Height          =   195
      Index           =   0
      Left            =   555
      TabIndex        =   0
      Top             =   255
      Width           =   270
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Coder       As cEncDec
Attribute Coder.VB_VarDescription = "Class instance"
Private Encrypted   As String
Attribute Encrypted.VB_VarDescription = "Encrypted string, used because the textboxes don't like Chr$(0)"
Private Decrypted   As String

Private Sub AdjUI(Enable As Boolean)

    btEncrypt.Enabled = Enable
    btDecrypt.Enabled = Enable
    txClear.Enabled = Enable
    txScrambled.Enabled = Enable
    Screen.MousePointer = IIf(Enable, vbDefault, vbHourglass)
    DoEvents

End Sub

Private Sub btDecrypt_Click()

    txClear = ""
    AdjUI False
    Err.Clear
    On Error Resume Next

      'decrypt text and compare signatures (from decrypted text and reconstructed)
      '======================================================================
      Decrypted = Coder.Decrypt(Encrypted, txKey)
      If Err = 0 Then 'no error from decryption, compare signatures
          If Left$(Decrypted, 16) = Coder.Signature(True, txKey & Mid$(Decrypted, 17), [Medium]) Then  'ok - display decrypted message
              txClear = Mid$(Decrypted, 17)
            Else 'NOT LEFT$(DECRYPTED,...
              txClear = ""
              Err.Raise 1002, , "The key is wrong."
          End If
      End If
      '======================================================================

      FeedMinMax Err
      AdjUI True
      If Err Then
          MsgBox "Cannot decrypt: " & Err.Description, , "Error " & Err & " from " & Err.Source
          txKey.SetFocus
      End If
    On Error GoTo 0

End Sub

Private Sub btEncrypt_Click()

  Dim Pointer

    txScrambled = ""
    txHex = ""
    AdjUI False
    Err.Clear
    On Error Resume Next

      'produce signature from key & text and then encrypt both the signature and the text
      '==============================================================================
      Encrypted = Coder.Encrypt(Coder.Signature(True, txKey & txClear, [Medium]) & txClear, txKey)
      '==============================================================================

      FeedMinMax Err
      AdjUI True
      If Err Then
          MsgBox "Cannot encrypt: " & Err.Description, , "Error " & Err & " from " & Err.Source
          txKey.SetFocus
          Encrypted = ""
      End If
    On Error GoTo 0
    txScrambled = Encrypted
    txScrambled.Refresh
    For Pointer = 1 To Len(Encrypted)
        On Error Resume Next
          txHex.SelText = Right$("0" & Hex$(Asc(Mid$(Encrypted, Pointer, 1))), 2) & " "
          If Err Then
              Exit For '>---> Next
          End If
        On Error GoTo 0
    Next Pointer
    txHex.SelStart = 0

End Sub

Private Sub FeedMinMax(HasNoValues)

    If HasNoValues Then
        lbMinMax = vbNullString
        lbCount = vbNullString
      Else 'HASNOVALUES = FALSE
        lbMinMax = "Min encrypted char length is " & Coder.MinCodeLength & ", max is " & Coder.MaxCodeLength & " bits"
        lbCount = Coder.BytesPerSecond & "b/s"
    End If

End Sub

Private Sub Form_Load()

    Set Coder = New cEncDec

End Sub

':) Ulli's VB Code Formatter V2.9.4 (06.02.2002 12:45:01) 6 + 98 = 104 Lines
