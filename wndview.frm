VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Post Message"
      Height          =   1815
      Left            =   4920
      TabIndex        =   22
      Top             =   4920
      Width           =   4215
      Begin VB.TextBox txtlParam 
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtWparam 
         Height          =   375
         Left            =   1800
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbMessages 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "lParam"
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "wParam"
         Height          =   375
         Left            =   720
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "WNDCLASS"
      Height          =   1335
      Left            =   4920
      TabIndex        =   17
      Top             =   3480
      Width           =   4215
      Begin VB.TextBox txtMenu 
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtStyle 
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Menu"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Style"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtWindowClass 
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtWindowText 
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Device Context "
      Height          =   1815
      Left            =   4920
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
      Begin VB.TextBox txtHdc 
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtClippingCapabilities 
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtColourPlanes 
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtBitsPerPixel 
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "hDC"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Clipping capabilities"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Colour planes"
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Bits per pixel"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CheckBox chkUnicode 
      Caption         =   "Unicode?"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtHwnd 
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   11668
      _Version        =   327682
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlWindows"
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Window Class"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Window Text"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Window Handle"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin ComctlLib.ImageList imlWindows 
      Left            =   2760
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wndview.frx":0000
            Key             =   "PARENT"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wndview.frx":0352
            Key             =   "CHILD"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'\\ Requires v1.1.3 or above of the EventVB dll
'\\ See http://www.merrioncomputing.com/Download/index.htm
Private WithEvents apiLink As EventVB.APIFunctions
Attribute apiLink.VB_VarHelpID = -1
Private AllWindows As Collection


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Sub AddChildren(wndThis As ApiWindow)

Dim wndChild As ApiWindow

For Each wndChild In wndThis.ChildWindows
    AllWindows.Add wndChild, "HWND:" & wndChild.hwnd
    With wndChild
        TreeView1.Nodes.Add "HWND:" & wndThis.hwnd, tvwChild, "HWND:" & .hwnd, "(" & .hwnd & ")" & .ClassName, "CHILD"
        Call AddChildren(wndChild)
    End With
Next wndChild

End Sub

Private Sub FillMessageList()

Me.cmbMessages.Clear

Dim msg As EventVB.WindowMessages

For msg = &H0 To &H500
    If InStr(apiLink.sGetMessageName(msg), "UNKNOWN") = 0 Then
        cmbMessages.AddItem apiLink.sGetMessageName(msg)
        cmbMessages.ItemData(cmbMessages.NewIndex) = msg
    End If
    
Next msg

End Sub

Private Property Let SelectedWindow(ByVal wndSelected As ApiWindow)

With wndSelected
    txtHwnd = .hwnd
    chkUnicode.Value = IIf(.Unicode, vbChecked, vbUnchecked)
    txtWindowText = .WindowText
    txtWindowClass = .ClassName
    With .DeviceContext
        txtHdc = .hDC
        txtBitsPerPixel = .BitsPerPixel
        txtColourPlanes = .ColourPlanes
        If .ClipingCapabilities = CP_NONE Then
            txtClippingCapabilities = "NONE"
        ElseIf .ClipingCapabilities = CP_RECTANGLE Then
            txtClippingCapabilities = "RECTANGLE"
        Else
            txtClippingCapabilities = "REGION"
        End If
    End With
    If Not (.WndClass Is Nothing) Then
    With .WndClass
        txtStyle = .Style
        txtMenu = .lpszMenuName
        
    End With
    End If
End With

End Property

Private Sub apiLink_ApiError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

MsgBox Description, vbCritical + vbOKOnly, "Error in " & Source

End Sub


Private Sub cmdPost_Click()

Dim lRet As Long

If Me.cmbMessages.ListIndex > 0 Then
    lRet = SendMessage(txtHwnd, cmbMessages.ItemData(cmbMessages.ListIndex), txtWparam, txtlParam)
End If

End Sub

Private Sub Form_Load()

Dim wndThis As EventVB.ApiWindow

Set AllWindows = New Collection
Set apiLink = New APIFunctions

For Each wndThis In apiLink.System.TopLevelWindows
    AllWindows.Add wndThis, "HWND:" & wndThis.hwnd
    With wndThis
        TreeView1.Nodes.Add , , "HWND:" & .hwnd, "(" & .hwnd & ")" & .ClassName, "PARENT"
        Call AddChildren(wndThis)
    End With
Next wndThis

Call FillMessageList

End Sub




Private Sub TreeView1_Click()

SelectedWindow = AllWindows.Item(TreeView1.SelectedItem.Key)


End Sub


Private Sub txtHwnd_Change()

If IsNumeric(txtHwnd) Then
    Me.cmdPost.Enabled = True
Else
    Me.cmdPost.Enabled = False
End If

End Sub


