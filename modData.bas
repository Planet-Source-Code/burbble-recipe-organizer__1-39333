Attribute VB_Name = "modData"
Type RecipeType
Name As String
Ingredients As String
Directions As String
End Type

Public Recipe() As RecipeType
Public RecipeAmount As Long

Public CurrentRecipe As Long

Sub LoadRecipes()
Dim tmpNothing As String

RecipeAmount = 0
Open App.Path & "\Recipes.dat" For Input As #10
Do While Not EOF(10)
Line Input #10, tmpNothing
RecipeAmount = RecipeAmount + 1
Loop
Close #10

RecipeAmount = (RecipeAmount - 1) / 3
If RecipeAmount = 0 Then
ReDim Recipe(0)
Exit Sub
End If

ReDim Recipe(RecipeAmount - 1)

Open App.Path & "\Recipes.dat" For Input As #10

For a = 0 To UBound(Recipe)
Line Input #10, tmpNothing
Recipe(a).Name = tmpNothing
Line Input #10, tmpNothing
Recipe(a).Ingredients = tmpNothing
Line Input #10, tmpNothing
Recipe(a).Directions = tmpNothing
Next a

Close #10
End Sub

Sub LoadRecipesPlus1(NewName As String, NewIngredients As String, NewDirections As String)
Dim tmpNothing As String

RecipeAmount = 0
Open App.Path & "\Recipes.dat" For Input As #10
Do While Not EOF(10)
Line Input #10, tmpNothing
RecipeAmount = RecipeAmount + 1
Loop
Close #10

RecipeAmount = (RecipeAmount - 1) / 3

ReDim Recipe(RecipeAmount)

Open App.Path & "\Recipes.dat" For Input As #10

For a = 0 To UBound(Recipe) - 1
Line Input #10, tmpNothing
Recipe(a).Name = tmpNothing
Line Input #10, tmpNothing
Recipe(a).Ingredients = tmpNothing
Line Input #10, tmpNothing
Recipe(a).Directions = tmpNothing
Next a

Close #10

Recipe(UBound(Recipe)).Name = NewName
Recipe(UBound(Recipe)).Ingredients = NewIngredients
Recipe(UBound(Recipe)).Directions = NewDirections

SaveRecipes
End Sub

Sub LoadRecipesMinus1(RecipeToDelete As Long)
LoadRecipes
Kill App.Path & "\Recipes.dat"
Open App.Path & "\Recipes.dat" For Output As #10
For h = 0 To UBound(Recipe)
If h = RecipeToDelete Then GoTo Next1
Print #10, Recipe(h).Name
Print #10, Recipe(h).Ingredients
Print #10, Recipe(h).Directions
GoTo Next1
Next1:
Next h
Close #10
End Sub

Sub SaveRecipes()
Kill App.Path & "\Recipes.dat"
Open App.Path & "\Recipes.dat" For Output As #10
For h = 0 To UBound(Recipe)
Print #10, Recipe(h).Name
Print #10, Recipe(h).Ingredients
Print #10, Recipe(h).Directions
Next h
Close #10
End Sub

Function GetRecipeNumber(RecipeName As String) As Long
For b = 0 To UBound(Recipe)
If Recipe(b).Name = RecipeName Then
GetRecipeNumber = b
Exit Function
End If
Next b
End Function
