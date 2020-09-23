VERSION 5.00
Begin VB.Form frmoopvb 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OOP and VB! by J. Brandon George"
   ClientHeight    =   3825
   ClientLeft      =   5160
   ClientTop       =   3690
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCol 
      Caption         =   "Get Collection"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdColCount 
      Caption         =   "Count Data Col."
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveasCol 
      Caption         =   "Save Data"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set Obj values"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Obj values"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "goto: http://www.xssi.net for more on OOP and VB!"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Date1:"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Num1:"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Info1:"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmoopvb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public obj As New clsObj
Public ColO As New ColObj
Private Sub cmdClear_Click()
  
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  
End Sub

Private Sub cmdCol_Click()
  Dim x
  
  Set x = ColO.Item(1)
  MsgBox "" & x.info1 & " " & x.num1 & " " & x.date1 & ""
End Sub

Private Sub cmdColCount_Click()
 MsgBox "" & ColO.Count & ""
 
End Sub


Private Sub cmdGet_Click()
  ' used to get the Values of the Class.Property named:info1, num1, and date1. <- this uses the Get methods
  Text1.Text = obj.info1
  Text2.Text = obj.num1
  Text3.Text = obj.date1
  
End Sub

Private Sub cmdSaveasCol_Click()
ColO.Add obj.info1, obj.num1, obj.date1
    
End Sub

Private Sub cmdSet_Click()
  ' used to set the values of the Class.Property named: info1, num1, and date1. <- this use's the Let methods
 obj.info1 = Text1.Text
 obj.num1 = Text2.Text
 obj.date1 = Text3.Text
 
End Sub



Private Sub Form_Load()
MsgBox "type some text in the Text box area's, then click the set obj values button. The clear the form and press the get. The code is very clean and should explain a little on the concepts behind OOP and VB. look for more soon!", vbInformation, "OOP and VB!"
End Sub
