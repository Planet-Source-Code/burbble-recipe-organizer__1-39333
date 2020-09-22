VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Burbble's Recipe Organizer"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic1 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4575
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   840
      Width           =   5655
      Begin VB.TextBox txtSearch1 
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   3600
         Width           =   2775
      End
      Begin VB.ComboBox cmbDisplay1 
         Height          =   330
         ItemData        =   "frmMain.frx":0442
         Left            =   2880
         List            =   "frmMain.frx":0497
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CommandButton cmdSearch1 
         Caption         =   "Search"
         Height          =   495
         Left            =   2880
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtFind1 
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton cmdView1 
         Caption         =   "View Recipe"
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.ListBox lstRecipes1 
         Height          =   4260
         ItemData        =   "frmMain.frx":04F5
         Left            =   0
         List            =   "frmMain.frx":04F7
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Use this to search recipe text for a word or phrase."
         Height          =   615
         Left            =   4320
         TabIndex        =   17
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Select a letter above to display only recipes beginning with that letter."
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Recipes:"
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Type in the name of a recipe, or the beginning of the name, above to highlight it in the list."
         Height          =   735
         Left            =   2880
         TabIndex        =   14
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Click View Recipe to see a recipe."
         Height          =   495
         Left            =   4320
         TabIndex        =   12
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Search within a Recipe for:"
         Height          =   495
         Left            =   2880
         TabIndex        =   10
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Display Recipes Beginning with:"
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Search for a Recipe Beginning with:"
         Height          =   615
         Left            =   2880
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.PictureBox pic2 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4575
      ScaleWidth      =   5655
      TabIndex        =   2
      Top             =   840
      Width           =   5655
      Begin VB.TextBox txtName2 
         Height          =   315
         Left            =   0
         TabIndex        =   22
         Top             =   240
         Width           =   5655
      End
      Begin VB.CommandButton cmdAdd2 
         Caption         =   "Add Recipe"
         Height          =   495
         Left            =   4320
         TabIndex        =   28
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtDirections2 
         Height          =   1455
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   2400
         Width           =   5655
      End
      Begin VB.TextBox txtIngredients2 
         Height          =   1215
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label12 
         Caption         =   "Recipe Name:"
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   5535
      End
      Begin VB.Label Label11 
         Caption         =   $"frmMain.frx":04F9
         Height          =   615
         Left            =   0
         TabIndex        =   20
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Label Label10 
         Caption         =   "Directions"
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   2160
         Width           =   5655
      End
      Begin VB.Label Label9 
         Caption         =   "Ingredients:"
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   600
         Width           =   5655
      End
   End
   Begin VB.PictureBox pic3 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4575
      ScaleWidth      =   5655
      TabIndex        =   3
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply Changes"
         Height          =   495
         Left            =   3600
         TabIndex        =   37
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtDirections3 
         Height          =   735
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtIngredients3 
         Height          =   735
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtName3 
         Height          =   315
         Left            =   2880
         TabIndex        =   32
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Delete Recipe"
         Height          =   495
         Left            =   2880
         TabIndex        =   27
         Top             =   0
         Width           =   1335
      End
      Begin VB.ListBox lstRecipes3 
         Height          =   4260
         ItemData        =   "frmMain.frx":05A5
         Left            =   0
         List            =   "frmMain.frx":05A7
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "Directions:"
         Height          =   375
         Left            =   2880
         TabIndex        =   35
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label17 
         Caption         =   "Ingredients:"
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Name:"
         Height          =   495
         Left            =   2880
         TabIndex        =   31
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label15 
         Caption         =   "Change portions of a recipe. Click Apply Changes when you are done."
         Height          =   495
         Left            =   2880
         TabIndex        =   30
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label Label14 
         Caption         =   "Permanently delete a recipe."
         Height          =   495
         Left            =   4320
         TabIndex        =   29
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Recipes:"
         Height          =   495
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2775
      End
   End
   Begin MSComctlLib.ImageList imglstMain 
      Left            =   5520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E51
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsMain 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9551
      HotTracking     =   -1  'True
      ImageList       =   "imglstMain"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "View Recipes"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Recipes"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit Recipes"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbDisplay1_Click()
If cmbDisplay1.Text = "Any letter" Then
LoadRecipes
lstRecipes1.Clear
For d = 0 To UBound(Recipe)
lstRecipes1.AddItem Recipe(d).Name
Next d
Else
LoadRecipes
lstRecipes1.Clear
For d = 0 To UBound(Recipe)
If LCase(Mid(Recipe(d).Name, 1, 1)) = LCase(cmbDisplay1.Text) Then
lstRecipes1.AddItem Recipe(d).Name
End If
Next d
End If
End Sub

Private Sub cmdAdd2_Click()
Dim tmpAdd1 As String
Dim tmpAdd2 As String
Dim tmpAdd3 As String

tmpAdd1 = txtName2.Text
tmpAdd2 = Replace(txtIngredients2.Text, Chr(13), "<enter>")
tmpAdd3 = Replace(txtDirections2.Text, Chr(13), "<enter>")
tmpAdd2 = Replace(tmpAdd2, Chr(10), "")
tmpAdd3 = Replace(tmpAdd3, Chr(10), "")

LoadRecipesPlus1 tmpAdd1, tmpAdd2, tmpAdd3

txtName2.Text = ""
txtIngredients2.Text = ""
txtDirections2.Text = ""

MsgBox "Added recipe successfully!", vbInformation
End Sub

Private Sub cmdApply_Click()
If lstRecipes3.ListIndex = -1 Then Exit Sub
If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
Dim tmpAdd1 As String
Dim tmpAdd2 As String
Dim tmpAdd3 As String

tmpAdd1 = txtName3.Text
tmpAdd2 = Replace(txtIngredients3.Text, Chr(13), "<enter>")
tmpAdd3 = Replace(txtDirections3.Text, Chr(13), "<enter>")
tmpAdd2 = Replace(tmpAdd2, Chr(10), "")
tmpAdd3 = Replace(tmpAdd3, Chr(10), "")

Recipe(GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))).Name = tmpAdd1
Recipe(GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))).Ingredients = tmpAdd2
Recipe(GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))).Directions = tmpAdd3
SaveRecipes
tsMain_Click
MsgBox "Successfully changed recipe.", vbInformation
End Sub

Private Sub cmdRemove_Click()
If lstRecipes3.ListIndex = -1 Then Exit Sub
If MsgBox("Are you sure? The recipe will be permanently deleted.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
LoadRecipesMinus1 GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))
tsMain_Click
MsgBox "Successfully deleted recipe.", vbInformation
End Sub

Private Sub cmdSearch1_Click()
If txtSearch1.Text = "" Then Exit Sub
LoadRecipes
lstRecipes1.Clear
For d = 0 To UBound(Recipe)
If InStr(1, LCase(Recipe(d).Directions), LCase(txtSearch1.Text)) <> 0 Then
lstRecipes1.AddItem Recipe(d).Name
ElseIf InStr(1, LCase(Recipe(d).Ingredients), LCase(txtSearch1.Text)) <> 0 Then
lstRecipes1.AddItem Recipe(d).Name
End If
Next d
End Sub

Private Sub cmdView1_Click()
If lstRecipes1.ListIndex = -1 Then Exit Sub
CurrentRecipe = GetRecipeNumber(lstRecipes1.List(lstRecipes1.ListIndex))
frmView.Show vbModal
End Sub

Private Sub Form_Load()
tsMain_Click
End Sub

Private Sub lstRecipes1_DblClick()
cmdView1_Click
End Sub

Private Sub lstRecipes3_Click()
If lstRecipes3.ListIndex = -1 Then Exit Sub
txtName3.Text = Recipe(GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))).Name
txtIngredients3.Text = Replace(Recipe(GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))).Ingredients, "<enter>", vbNewLine)
txtDirections3.Text = Replace(Recipe(GetRecipeNumber(lstRecipes3.List(lstRecipes3.ListIndex))).Directions, "<enter>", vbNewLine)
End Sub

Private Sub tsMain_Click()
If tsMain.SelectedItem.Index = 1 Then
pic1.Visible = True
pic2.Visible = False
pic3.Visible = False
LoadRecipes
lstRecipes1.Clear
For d = 0 To UBound(Recipe)
lstRecipes1.AddItem Recipe(d).Name
Next d
cmbDisplay1.ListIndex = 0
ElseIf tsMain.SelectedItem.Index = 2 Then
pic1.Visible = False
pic2.Visible = True
pic3.Visible = False
ElseIf tsMain.SelectedItem.Index = 3 Then
LoadRecipes
lstRecipes3.Clear
For d = 0 To UBound(Recipe)
lstRecipes3.AddItem Recipe(d).Name
Next d
pic1.Visible = False
pic2.Visible = False
pic3.Visible = True
End If
End Sub

Private Sub txtFind1_Change()
If txtFind1.Text = "" Then Exit Sub
For c = 0 To lstRecipes1.ListCount - 1
If Left(lstRecipes1.List(c), Len(txtFind1.Text)) = txtFind1.Text Then
lstRecipes1.ListIndex = c
Exit Sub
End If
Next c
End Sub
