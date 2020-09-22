VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Viewing Recipe"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtDirections 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2040
      Width           =   5655
   End
   Begin VB.ListBox lstIngredients 
      Appearance      =   0  'Flat
      Height          =   1230
      ItemData        =   "frmView.frx":0000
      Left            =   120
      List            =   "frmView.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblPrint 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label2 
      Caption         =   "Directions:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Ingredients:"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim tmpPrint As String
tmpPrint = ""
Printer.CurrentX = 120
Printer.CurrentY = 120
Printer.Font = "Arial"
Printer.FontSize = 18
Printer.FontBold = True
Printer.Print Recipe(CurrentRecipe).Name
Printer.FontSize = 12
Printer.Print vbNewLine & vbNewLine & "Ingredients:" & vbNewLine
Printer.FontBold = False
For f = 0 To lstIngredients.ListCount - 1
Printer.Print lstIngredients.List(f)
Next f
Printer.FontBold = True
Printer.Print vbNewLine & vbNewLine & "Directions:" & vbNewLine
Printer.FontBold = False

For f = 1 To Len(txtDirections.Text)

If Mid(txtDirections.Text, f, 1) = " " And lblPrint.Width >= 10800 Then
tmpPrint = tmpPrint & lblPrint.Caption & vbNewLine
lblPrint.Caption = ""
ElseIf lblPrint.Width >= 12000 Then
tmpPrint = tmpPrint & lblPrint.Caption & Mid(txtDirections.Text, f, 1) & vbNewLine
lblPrint.Caption = ""
ElseIf Mid(txtDirections.Text, f, 1) <> " " And lblPrint.Width >= 10800 Then
lblPrint.Caption = lblPrint.Caption & Mid(txtDirections.Text, f, 1)
ElseIf lblPrint.Width < 10800 Then
lblPrint.Caption = lblPrint.Caption & Mid(txtDirections.Text, f, 1)
End If

If f = Len(txtDirections.Text) Then
tmpPrint = tmpPrint & lblPrint.Caption & vbNewLine
lblPrint.Caption = ""
End If
Next f

Printer.Print tmpPrint
Printer.EndDoc
MsgBox "Sent recipe to printer.", vbInformation
End Sub

Private Sub Form_Load()
Dim tmpIngredients As String

lstIngredients.Clear
txtDirections.Text = Replace(Recipe(CurrentRecipe).Directions, "<enter>", vbNewLine)

frmView.Caption = "Viewing Recipe - " & Recipe(CurrentRecipe).Name

txtTemp.Text = Replace(Recipe(CurrentRecipe).Ingredients, "<enter>", Chr(13))

For e = 1 To Len(txtTemp.Text)
If Mid(txtTemp.Text, e, 1) = Chr(13) Then
lstIngredients.AddItem tmpIngredients
tmpIngredients = ""
Else
tmpIngredients = tmpIngredients & Mid(txtTemp.Text, e, 1)
End If
Next e
lstIngredients.AddItem tmpIngredients
lstIngredients.ListIndex = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmView = Nothing
End Sub
