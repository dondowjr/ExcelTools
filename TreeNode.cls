VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
End
Attribute VB_Name = "TreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit On

' © 2024 Donald R Dow Jr, All Rights Reserved.


' TreeNode Class
' --------------
' The TreeNode class is a generic class that stores various
' types of data. The TreeNode can contain unlimited number
' of children TreeNodes.


' Code organization/outline
' -------------------------
'   + Module level declaration
'   + Init and load methods
'   + Save methods
'   + Properties
'   + Public methods/functions
'   + Public helper methods/functions
'   + Private helper methods/functions


' Naming conventions
' ------------------
' This macro used the older VB naming convention that was
' established by Microsoft in the Visual Basic v3 - v6
' era.
'
' Variable names use title casing with a lower-case scope
' prefix character. Prefixes used here are:
'
'  m = module or class level
'  v = method level
'  c = constant

' Revision Notes
' --------------
' 6/18/2024 DRD Created class


' enums
Public Enum TNInfoEnum
    tnName = 0
    tnTag = 1
    tnNumber = 2
    tnText = 3
End Enum

' objects
Private mItems As Collection
Private mObject As Object
Private mCache As TreeNode

' numbers
Private mNumber As Double

' strings
Private mTag As String
Private mName As String
Private mText As String


' -----------------------------------------------------------
' Load and save methods
' -----------------------------------------------------------

Public Sub Init(ByVal vName As String,
                Optional ByVal vTag As String,
                Optional ByVal vNumber As Double,
                Optional ByVal vText As String)

    mName = vName
    mNumber = vNumber
    mTag = vTag
    mText = vText

End Sub


Public Sub LoadFromString(ByVal vData As String)

    ' Expecting data to be in format of name:tag:number:text

    Dim vParts() As String
    Dim i As Long

    vParts = Split(vData, ":")

    For i = 0 To UBound(vParts)
        vData = Trim(vParts(i))
        Select Case i
            Case 0
                mName = vData
            Case 1
                mTag = vData
            Case 2
                mNumber = vData
            Case 3
                mText = vData
        End Select
    Next

End Sub


Public Function SaveToString() As String

    ' Expecting data to be in format of:
    '   name:tag:number:text

    Dim vResults As String

    vResults = mName & ":" &
               mTag & ":" &
               mNumber & ":" &
               mText

    SaveToString = vResults

End Function


Public Property Get Count() As Long

' Returns a count of children attached to the current node

If Not mItems Is Nothing Then
Count = mItems.Count
End If

End Property


' -----------------------------------------------------------
' Properties
' -----------------------------------------------------------


Public Property Get Name() As String

    ' Returns the name of the node.
    
    ' This is a read only property as the name is typically
    '  used as the index in the parent collection.
    
    Name = mName
    
End Property


Public Property Get Number() As Double

    Number = mNumber

End Property
Public Property Let Number(ByVal vData As Double)

    mNumber = vData

End Property


Public Property Get Tag() As String

    Tag = mTag
    
End Property
Public Property Let Tag(ByVal vData As String)

    mTag = vData
    
End Property


Public Property Get Text() As String

    Text = mText
    
End Property
Public Property Let Text(ByVal vText As String)

    mText = vText
    
End Property



' Notes: 1. The name of this property is abbreviated
'           as Object is a reserved word.
'        2. The obj property is not cloned when using
'           the Clone function.

Public Property Get Obj() As Object

    Set Obj = mObject

End Property
Public Property Set Obj(ByVal vObject As Object)

    Set mObject = vObject

End Property


Public Property Get Item(ByVal vName As String) As TreeNode

    ' Returns an child node based on vName.
    
    ' Note: The if a child node with vName does not exist,
    '       one will be created.

    InitItems
    
    If Exists(vName) Then
Set Item = mCache
    Else
Set Item = NewNode(vName)
    End If

End Property


Public Property Get Items() As Collection
    
    ' Returns the generic collections object containing the
    ' children nodes of the current node.
    
    InitItems
    
    Set Items = mItems
    
End Property


' -----------------------------------------------------------
' Public methods/functions
' -----------------------------------------------------------


Public Sub AddNumber(ByVal vNumber As Double)

    ' Adds the vNumber to the existing number value.
    '
    ' This method is typically used to accumulate a value.

    mNumber = mNumber + vNumber

End Sub


Public Function NewNode(ByVal vName As String,
                        Optional ByVal vTag As String,
                        Optional ByVal vNumber As Double,
                        Optional ByVal vText As String) As TreeNode

    ' Creates a new child node with the given properties, adds
    ' it to the collection and returns the newly created node.
    '
    ' If the node already exists, the node will be updated and
    ' a reference returned to that node.

    InitItems()

    If Exists(vName) Then
        mCache.Init vName, vTag, vNumber, vText
    Else
        Set mCache = New TreeNode
        mCache.Init vName, vTag, vNumber, vText
        mItems.Add mCache, vName
    End If
    
    Set NewNode = mCache
    
End Function


Public Sub AddNode(ByRef vNode As TreeNode)

    ' Adds a new child node to the current node.
    '
    ' If a child node with the same name already exists,
    ' the existing child node will be updated with the
    ' vNode values. The vNode argument will then be
    ' set to the node object in the collection.

    InitItems()

    If Exists(vNode.Name) Then
        With vNode
            mCache.Init.Name, .Tag, .Number, .Text
        End With
        Set vNode = mCache
    Else
        mItems.Add vNode, vNode.Name
    End If

End Sub


Public Sub AddChildren(ByVal vData As String,
                       Optional vChildDelim = ",",
                       Optional vInfoDelim = ":")

    ' Adds multiple children to the current node.
    '
    ' vChildDelim = delimiter separating the child node data blocks.
    ' vInfoDelim  = delimiter separating the values within a a data block.


    Dim vParts() As String
    Dim vItem As TreeNode
    Dim i As Long

    vParts = Split(vData, vChildDelim)

    For i = 0 To UBound(vParts)
        Set vItem = New TreeNode
        vItem.Load Trim(vParts(i))
        AddNode vItem
    Next

End Sub


Public Function Exists(ByVal vName As String) As Boolean

    ' Returns true a child node of vName exists.

    ' If the node does exist. Then the module level cache
    ' object is set. This is done so additional fetches for
    ' a child of the same node is performant.

    On Error Resume Next

    InitItems()

    If Not mCache Is Nothing Then

        If mCache.Name = vName Then
            Exists = True
            Exit Function
        Else
            Set mCache = Nothing
        End If

    End If
    
    Set mCache = mItems.Item(vName)
        
    Exists = Not mCache Is Nothing

    Err.Clear

End Function


Public Function NestedItem(ParamArray vName() As Variant) As TreeNode

    ' Returns a nested node.

    ' Example NestedItem("Level1", "Level2", "Level3" [, ...])

    Dim vItem As TreeNode
    Dim vTemp As TreeNode
    Dim i As Long
    
    Set vItem = Item(vName(0))

    For i = 1 To UBound(vName)
        Set vTemp = vItem.Item(vName(i))
        Set vItem = vTemp
    Next

    Set NestedItem = vItem

End Function


Public Function ChildValues(ByVal vInfo As TNInfoEnum,
                            Optional ByVal vDelim As String = ",") As String

    ' Returns a formatted string of all the child values.

    ' vInfo  = Determines what value will be returned.
    ' vDelim = Is used to delimit the values.

    Dim i As Long
    Dim vResult As String
    Dim vValue As String
    Dim vNumber As Double

    For i = 1 To Count

        If vInfo = tnNumber Then
            vNumber = vNumber + mItems(i).Number
        Else
            vResult = StrAppend(vResult, mItems(i).Value(vInfo), vDelim)
        End If

    Next

    If vInfo = tnNumber Then
        ChildValues = vNumber
    Else
        ChildValues = vResult
    End If

End Function


Public Function Clone(Optional ByVal vIncChildren As Boolean) As TreeNode

    ' Returns a clone of the current node.

    ' vIncChildren = If true, clones of the children to added to the
    '                the returned node.
    '
    ' Note: The obj property is NOT cloned.

    Dim vItem As TreeNode
    Dim i As Long
    
    Set vItem = New TreeNode
    
    vItem.Init mName, mTag, mNumber, mText

    If vIncChildren Then

        For i = 1 To Count

            AddNode mItems(i).Clone(vIncChildren)

        Next

    End If
    
    Set Clone = vItem

End Function


Public Function Value(ByVal vInfo As TNInfoEnum) As String

    ' Returns a string value of a property given vInfo.

    Dim vValue As String

    Select Case vInfo
        Case tnName
            vValue = mName
        Case tnTag
            vValue = mTag
        Case tnNumber
            vValue = mNumber
        Case tnText
            vValue = mText
    End Select

    Value = vValue

End Function


Public Function DumpTree(Optional ByVal vIndent As Long = 1) As String

    ' Returns a formatted string of the current TreeNode and
    ' child nodes.

    Dim i As Long
    Dim vValue As String

    If vValue <= 0 Then vValue = 1

    vValue = String(vIndent * 2, " ") & mName & ":" & mTag & ":" & mNumber & ":" & mText

    For i = 1 To Count
        vValue = StrAppend(vValue, mItems.Item(i).DumpTree(vIndent + 1), vbCrLf)
    Next

    DumpTree = vValue

End Function


Public Property Get This() As TreeNode

    ' Returns the current object.

    ' This is handy when a variable is used within a With block
    ' and you need to get a passable reference the object.

    Set This = Me

End Property


Public Function GetNewSerialName(ByVal vNodeName As String) As String

    ' Returns a serialized child name that can be used for adding
    ' a new child to the collection.

    ' Serialization is done by adding an incremental number at the
    ' end of the vNodeName value. If the vNodeName is blank, the name
    ' of "Item" is used.

    Dim vTemp As String
    Dim c As Long

    ' add a name if the node name is blank
    If LenB(vNodeName) = 0 Then
        vNodeName = "Item"
    End If

    ' serialize the node name if needed
    If Exists(vNodeName) Then

        Do
            c = c + 1
            vTemp = vNodeName & c
        Loop Until Not Exists(vTemp)

        vNodeName = vTemp

    End If

    GetNewSerialName = vNodeName

End Function


Public Function SumChildren(Optional ByVal vRecurse = False) As Double

    ' Returns the sum of the number property of all the children in the
    ' current node.
    '
    ' vRecurse = If true, all children and sub-children will be totaled.
    '            If false, just the emidiate children of the parent are totaled.

    Dim vResult As Double
    Dim vNode As TreeNode

    For Each vNode In mItems

        vResult = vResult + vNode.Number

        If vRecurse Then

            vResult = vResult + vNode.SumChildren(vRecurse)

        End If

    Next

    TreeNode_SumChildren = vResult

End Function

Public Function FindNode(ByVal vValueMask As String,
                         Optional ByVal vProp As TNInfoEnum = tnName,
                         Optional ByVal vAllNodes = False) As TreeNode

    ' Finds nodes based on a property search value.
    '
    ' vValueMask = The value of mask value used to qualify nodes. This
    '              value can contain "Like" wild card such as *, ?, ~
    '
    ' vProp      = The node property to search.
    '
    ' vAllNodes  = If false, then only the first node found is returned.
    '              If true, then all nodes are returned that match the
    '              search criteria. The node returned is a container of
    '              child nodes that were found.


End Function


' -----------------------------------------------------------
' Private methods/functions
' -----------------------------------------------------------


Private Sub InitItems()

    ' Inits mItems if it has not already been Init'ed.

    If mItems Is Nothing Then
        Set mItems = New Collection
    End If

End Sub


