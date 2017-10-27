VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As Integer
Public name As String

Private pUsers, pMemberships As Collection
Public Property Get memberships() As Collection
  Set memberships = pMemberships
End Property
Public Property Get users() As Collection
  Set users = pUsers
End Property

'Constructor
Private Sub Class_Initialize()
  Set pUsers = New Collection
  Set pMemberships = New Collection
End Sub
'Destructor
Private Sub Class_Terminate()
  Set pUsers = Nothing
  Set pMemberships = Nothing
End Sub

Public Function load(ByVal json As Dictionary)
  Me.id = Int(json("id"))
  Me.name = json("name")
  
  Dim val As Object
  If json.Exists("users") Then
    Dim u As RedmineUser
    For Each val In json("users")
      Set u = New RedmineUser
      u.id = Int(val("user")("id"))
      u.name = val("user")("name")
      pUsers.Add u
    Next
  End If
  
  If json.Exists("memberships") Then
    Dim p As RedmineProject
    For Each val In json("memberships")
      Set p = New RedmineProject
      p.id = Int(val("project")("id"))
      p.name = val("project")("name")
      pMemberships.Add p
    Next
  End If
End Function


