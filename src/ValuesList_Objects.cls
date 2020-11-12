VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesList_Objects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "IValuesList implementation for random value generation. Dependency of 'ArrayGenerator'"
'@Folder("BetterArray.Tests.Dependencies.IValuesList.Implementations")
'@ModuleDescription("IValuesList implementation for random value generation. Dependency of 'ArrayGenerator'")

Option Explicit
Implements IValuesList

Private IsObjectType As Boolean

Private Sub Class_Initialize()
    IsObjectType = True
End Sub

Public Property Get IValuesList_IsObjectType() As Boolean
    IValuesList_IsObjectType = IsObjectType
End Property

Public Function IValuesList_GetRandomValue() As Variant
    Dim Result As Collection
    Set Result = New Collection
    Result.Add Rand(-1000, 1000)
    Set IValuesList_GetRandomValue = Result
End Function

Private Function Rand(ByVal Max As Long, Optional ByVal Min As Long = 0) As Long
    Randomize: Rand = Int(Rnd * (Max - Min) + 1) + Min
End Function
