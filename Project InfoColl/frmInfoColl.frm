VERSION 5.00
Begin VB.Form frmInfoColl 
   Caption         =   "Hellbound: Pal Information Collector"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Room Information"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      Begin VB.CommandButton cmdGetInfo 
         Caption         =   "Get Info"
         Height          =   330
         Left            =   4320
         TabIndex        =   7
         Top             =   1035
         Width           =   1275
      End
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   6
         Top             =   720
         Width           =   4110
      End
      Begin VB.TextBox txtClass 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   5
         Top             =   450
         Width           =   4110
      End
      Begin VB.TextBox txthWnd 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1485
         TabIndex        =   4
         Top             =   180
         Width           =   4110
      End
      Begin VB.Label Label3 
         Caption         =   "Room Title"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Room Class"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   495
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Room Handle"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmInfoColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Once you have called the GetPalWindow function it fills 3 global
'variables called
'Global lngRhWnd As Long ' RoomHandle
'Global strWindowTitle As String ' Room Title
'Global strClassname As String ' Room ClassName
'this info is sufficent to do anything else u may wish too
'when i started out in paltalk programming this was the hardest thing to do

Private Sub cmdGetInfo_Click()
Call GetPalWindow
txthWnd.Text = lngRhWnd
txtClass.Text = strClassname
txtTitle.Text = strWindowTitle
End Sub

