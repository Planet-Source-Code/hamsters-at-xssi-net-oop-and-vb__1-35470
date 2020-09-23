VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "this is the first Class in our OOP example"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************
' Basics of OOP in VB!
' ****
' Author: J. Brandon George
' email: josephbg@aol.com
' *************************
' This is a VERY BASIC overview of how to use classes and OOP
' ***************************************


'local variable(s) to hold property value(s)
Private mvarinfo1 As String 'local copy
Private mvarnum1 As Integer 'local copy
Private mvardate1 As Date 'local copy


Public Property Let date1(ByVal vData As Date)
Attribute date1.VB_Description = "date1 is a date"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.date1 = 5
    mvardate1 = vData
End Property


Public Property Get date1() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.date1
    date1 = mvardate1
End Property



Public Property Let num1(ByVal vData As Integer)
Attribute num1.VB_Description = "num1 one is a interger"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.num1 = 5
    mvarnum1 = vData
End Property


Public Property Get num1() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.num1
    num1 = mvarnum1
End Property



Public Property Let info1(ByVal vData As String)
Attribute info1.VB_Description = "info1 is used to display how we would turn a regular String var. into a String Property"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.info1 = 5
    mvarinfo1 = vData
End Property


Public Property Get info1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.info1
    info1 = mvarinfo1
End Property



