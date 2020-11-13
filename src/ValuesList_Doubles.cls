VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuesList_Doubles"
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

Private Sub Class_Initialize()
    IsObjectType = False
End Sub

Public Property Get IValuesList_IsObjectType() As Boolean
    IValuesList_IsObjectType = IsObjectType
End Property

Public Function IValuesList_GetRandomValue() As Variant
    IValuesList_GetRandomValue = Rand(-1000, 1000) * Rnd
End Function

Private Function Rand(ByVal Max As Long, Optional ByVal Min As Long = 0) As Long
    Randomize: Rand = Int(Rnd * (Max - Min) + 1) + Min
End Function

