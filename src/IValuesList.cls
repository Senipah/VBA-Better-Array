VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValuesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface definition for random value generation. Dependency of 'ArrayGenerator'"
'@Folder("BetterArray.Tests.Dependencies.IValuesList.Implementations")
'@ModuleDescription("Interface definition for random value generation. Dependency of 'ArrayGenerator'")

Option Explicit

Public Property Get IsObjectType() As Boolean

End Property

Public Function GetRandomValue() As Variant

End Function
