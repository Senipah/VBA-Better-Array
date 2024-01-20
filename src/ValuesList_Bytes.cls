VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesList_Bytes"
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
Private Values(0 To 255) As Byte

Private Sub Class_Initialize()
    Dim i As Long
    For i = LBound(Values) To UBound(Values)
        Values(i) = i
    Next
    IsObjectType = False
End Sub

Public Property Get IValuesList_IsObjectType() As Boolean
    IValuesList_IsObjectType = IsObjectType
End Property

Public Function IValuesList_GetRandomValue() As Variant
    IValuesList_GetRandomValue = Values(Rand(LBound(Values), UBound(Values)))
End Function

Private Function Rand(ByVal Max As Long, Optional ByVal Min As Long = 0) As Long
    Randomize: Rand = Int(Rnd * (Max - Min) + 1) + Min
End Function
