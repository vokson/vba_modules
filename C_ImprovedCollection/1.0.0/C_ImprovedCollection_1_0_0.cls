VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ImprovedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'============================================
' Author: nerv
' Last Update: 25/07/2012
' E-mail: nerv-net@yandex.ru
'============================================
Option Explicit
Private this As Collection


Private Sub Class_Initialize()
    Set this = New Collection
End Sub


Private Sub Class_Terminate()
    Set this = Nothing
End Sub


'---------------------------------------------
' Name: Add Method
' Return: -
' Last Update: 21/07/2012
' About: It's the standard method
'---------------------------------------------
Public Sub Add(ByRef item As Variant, _
               Optional ByRef Key As Variant, _
               Optional ByRef Before As Variant, _
               Optional ByRef After As Variant)
    If IsMissing(Key) Then
        If Not IsMissing(Before) Then
            this.Add item:=item, Before:=Before
        ElseIf Not IsMissing(After) Then
            this.Add item:=item, After:=After
        Else
            this.Add item:=item
        End If
    Else
        If Not IsMissing(Before) Then
            this.Add item:=item, Key:=Key, Before:=Before
        ElseIf Not IsMissing(After) Then
            this.Add item:=item, Key:=Key, After:=After
        Else
            this.Add item:=item, Key:=Key
        End If
    End If
End Sub


'---------------------------------------------
' Name: Remove Method
' Return: removed value
' Last Update: 24/07/2012
' About: It's the standard method, but it returns the removed value
'---------------------------------------------
Public Function Remove(ByRef Index As Variant) As Variant
    If IsObject(this.item(Index)) Then
        Set Remove = this.item(Index)
    Else
        Remove = this.item(Index)
    End If
    this.Remove (Index)
End Function


'---------------------------------------------
' Name: RemoveAll Method
' Return: -
' Last Update: 21/07/2012
' About: It's the standard method. It's taken from the dictionary
'---------------------------------------------
Public Sub RemoveAll()
    Set this = New Collection
End Sub


'---------------------------------------------
' Name: Item Method Get
' Return: value
' Last Update: 24/07/2012
' About: It's the standard method
'---------------------------------------------
Public Property Get item(Optional ByRef Index As Variant) As Variant
Attribute item.VB_UserMemId = -4
    If IsMissing(Index) Then
        Set item = this.[_NewEnum]
    Else
        If IsObject(this.item(Index)) Then
            Set item = this.item(Index)
        Else
            item = this.item(Index)
        End If
    End If
End Property


'---------------------------------------------
' Name: Item Property
' Return: -
' Last Update: 21/07/2012
' About: It sets the value. If Index is String, then the item is added to the end
'---------------------------------------------
Public Property Let item(ByRef Index As Variant, _
                         ByRef Value As Variant)
Attribute item.VB_UserMemId = 0
    If VarType(Index) = vbString Then
        this.Remove (Index)
        this.Add item:=Value, Key:=Index
    Else
        this.Add item:=Value, After:=Index
        this.Remove (Index)
    End If
End Property


'---------------------------------------------
' Name: Count Method
' Return: number of items
' Last Update: 21/07/2012
' About: It's the standard method
'---------------------------------------------
Public Property Get Count() As Long
    Count = this.Count()
End Property


'---------------------------------------------
' Name: Exists Method
' Return: true / false
' Last Update: 22/07/2012
' About: It's checks item on the key
'---------------------------------------------
Public Function Exists(ByRef Key As String) As Boolean
    On Error Resume Next
    Exists = Not IsEmpty(this.item(Key))
End Function


'---------------------------------------------
' Name: ToString Method
' Return: string
' Last Update: 24/07/2012
' About: It's creates string from the collection
'---------------------------------------------
Public Property Get ToString(Optional ByRef Separator As String = " ") As String
    Dim arr() As Variant, col As New Collection, i As Long
    If this.Count > 0 Then
        Call AddItems(this, col)
        If col.Count > 0 Then
            ReDim arr(1 To col.Count)
            For i = col.Count To 1 Step -1
                arr(i) = col.item(i)
                col.Remove (i)
            Next
            Set col = Nothing
        End If
        ToString = Join(arr, Separator)
        Erase arr
    End If
End Property


'---------------------------------------------
' Name: For ToString Method
' Return: -
' Last Update: 24/07/2012
' About: It's recursive function. It adds items to the collection
'---------------------------------------------
Private Sub AddItems(ByRef that As Variant, ByRef col As Collection)
    Dim elem As Variant
    For Each elem In that
        If IsObject(elem) Then
            If elem Is Nothing Then
                col.Add "Nothing"
            Else
                If TypeOf elem Is Collection Or _
                   TypeOf elem Is CImprovedCollection Then
                    Call AddItems(elem, col)
                End If
            End If
        Else
            If IsArray(elem) Then
                Call AddItems(elem, col)
            Else
                col.Add elem
            End If
        End If
    Next
End Sub


'---------------------------------------------
' Name: ToArray Method
' Return: array
' Last Update: 24/07/2012
' About: It's creates array from the collection
'---------------------------------------------
' reserved: OnlyPrimitives
Public Property Get ToArray(Optional ByRef Dimension As Byte = 2) As Variant
    Dim arr() As Variant, i As Long
    If this.Count > 0 Then
        If Dimension = 2 Then
            ReDim arr(1 To this.Count, 1 To 1)
            For i = 1 To this.Count
                If IsObject(this.item(i)) Then
                    Set arr(i, 1) = this.item(i)
                Else
                    arr(i, 1) = this.item(i)
                End If
            Next
        ElseIf Dimension = 1 Then
            ReDim arr(1 To this.Count)
            For i = 1 To this.Count
                If IsObject(this.item(i)) Then
                    Set arr(i) = this.item(i)
                Else
                    arr(i) = this.item(i)
                End If
            Next
        End If
        ToArray = arr
        Erase arr
    End If
End Property


'---------------------------------------------
' Name: AddUnique Method
' Return: -
' Last Update: 25/07/2012
' About: It's adds to the collection only unique records
'---------------------------------------------
Public Sub AddUnique(ByRef data As Variant)
    If IsObject(data) Then
        If data Is Nothing Then
            Call Unique(Array("Nothing"))
        Else
            If TypeOf data Is Range Then
                If data.Cells.Count = 1 Then
                    Call Unique(Array(data.Value))
                Else
                    Call Unique(data.Value)
                End If
            ElseIf TypeOf data Is Collection Or _
                   TypeOf data Is CImprovedCollection Then
                Call Unique(data)
            End If
        End If
    Else
        If Not IsArray(data) Then
            Call Unique(Array(data))
        Else
            Call Unique(data)
        End If
    End If
End Sub


'---------------------------------------------
' Name: For AddUnique Method
' Return: -
' Last Update: 25/07/2012
' About: It's creates unique collection
'---------------------------------------------
Private Sub Unique(ByRef data As Variant)
    Dim elem As Variant
    On Error Resume Next
    For Each elem In data
        If IsObject(elem) Then
            Call Me.AddUnique(elem)
        ElseIf IsArray(elem) Then
            Call Unique(elem)
        Else
            If Not IsEmpty(elem) And Not IsError(elem) Then
                this.Add item:=elem, Key:=CStr(elem)
            End If
        End If
    Next
End Sub


'---------------------------------------------
' Name: Sort Method
' Return: -
' Last Update: 24/07/2012
' About: It's exchange sort
'---------------------------------------------
Public Sub Sort(Optional ByRef Order As Byte = xlAscending) ' 1
    Dim temp As Variant, outer As Long, inner As Long, min As Long
    For outer = 1 To this.Count - 1
        min = outer
        For inner = outer + 1 To this.Count
            If Not IsArray(this.item(min)) And _
               Not IsArray(this.item(inner)) And _
               Not IsObject(this.item(min)) And _
               Not IsObject(this.item(inner)) Then
                If Order = 1 Then
                    If this.item(inner) < this.item(min) Then
                        min = inner
                    End If
                ElseIf Order = 2 Then
                    If this.item(inner) > this.item(min) Then
                        min = inner
                    End If
                End If
            End If
        Next
        If min > outer Then
            temp = this.item(outer)
            Me.item(outer) = this.item(min)
            Me.item(min) = temp
        End If
    Next
End Sub
