VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_LambdaFunctionWizard 
   Caption         =   "Insert Power Function"
   ClientHeight    =   6530
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5170
   OleObjectBlob   =   "uf_LambdaFunctionWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_LambdaFunctionWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypePowerFunctionWizard
    LambdaStorage As zLIB_ListStorage
    EventsAreEnabled As Boolean
End Type
Private this As TypePowerFunctionWizard


Property Set LambdaStorage(ByRef storage)
'cannot pass variables to  userform event so store as a class property (userforms are classes)
    Set this.LambdaStorage = storage
End Property


Property Let EnableEvents(Enable As Boolean)

End Property


Property Let LambdaCategories(ByRef Categories)
    
    Dim i As Integer
    
    Me.comboCategories.AddItem "All"
    For i = LBound(Categories) To UBound(Categories)
        Me.comboCategories.AddItem Categories(i)
    Next i
    Me.comboCategories.Value = "All"

End Property



Private Sub comboCategories_Change()

    Dim LambdaNamesPerCategorySelection
    Dim i As Integer
    
    ReadLambdaNamesPerCategory this.LambdaStorage, LambdaNamesPerCategorySelection, Me.comboCategories.Value
    
    Me.lbFunctions.Clear
    For i = LBound(LambdaNamesPerCategorySelection) To UBound(LambdaNamesPerCategorySelection)
        Me.lbFunctions.AddItem LambdaNamesPerCategorySelection(i)
    Next i

End Sub




Private Sub UserForm_Terminate()
   Set this.LambdaStorage = Nothing
End Sub
