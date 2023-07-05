VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B_Versandstelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public gruppe As Integer
Public lieferungen As Collection

Private Sub Class_Initialize()
    Set lieferungen = New Collection
End Sub

