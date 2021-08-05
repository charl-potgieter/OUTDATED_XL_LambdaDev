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

Private Sub comboCategories_Change()
    MsgBox Me.comboCategories.Value
End Sub
