VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesList_Variants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "IValuesList implementation for random value generation. Dependency of 'ArrayGenerator'"
'@Folder("VBABetterArray.Tests.Dependencies.IValuesList.Implementations")
'@ModuleDescription("IValuesList implementation for random value generation. Dependency of 'ArrayGenerator'")

Option Explicit
Implements IValuesList

Private IsObjectType As Boolean
Private Const VALUE_COUNT As Long = 100
Private Values(VALUE_COUNT) As Variant

Private Sub Class_Initialize()
    Dim TypesList(0 To 4) As IValuesList
    Set TypesList(0) = New ValuesList_Booleans
    Set TypesList(1) = New ValuesList_Bytes
    Set TypesList(2) = New ValuesList_Doubles
    Set TypesList(3) = New ValuesList_Longs
    Set TypesList(4) = New ValuesList_Strings
    
    Dim i As Long
    For i = 0 To VALUE_COUNT
        Values(i) = TypesList(Rand(4)).GetRandomValue
    Next
    IsObjectType = False
End Sub

Public Property Get IValuesList_IsObjectType() As Boolean
    IValuesList_IsObjectType = IsObjectType
End Property

Private Function IValuesList_GetRandomValue() As Variant
    IValuesList_GetRandomValue = Values(Rand(LBound(Values), UBound(Values)))
End Function

Private Function Rand(ByVal Max As Long, Optional ByVal Min As Long = 0) As Long
    Randomize: Rand = Int(Rnd * (Max - Min) + 1) + Min
End Function

