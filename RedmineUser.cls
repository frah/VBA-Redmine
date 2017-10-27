VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id, status As Integer
Public login, firstname, lastname, mail As String

Private pMemberships, pGroups As Collection
Public Property Get memberships() As Collection
  Set memberships = pMemberships
End Property
Public Property Get groups() As Collection
  Set groups = pGroups
End Property

'Constructor
Private Sub Class_Initialize()
  Set pMemberships = New Collection
  Set pGroups = New Collection
End Sub
'Destructor
Private Sub Class_Terminate()
  Set pMemberships = Nothing
  Set pGroups = Nothing
End Sub

Public Function load(ByVal json As Dictionary)
  Me.id = Int(json("id"))
  Me.login = json("login")
  Me.firstname = json("firstname")
  Me.lastname = json("lastname")
  Me.mail = json("mail")
  Me.status = Int(json("status"))
  
  Dim val As Object
  If json.Exists("memberships") Then
    Dim p As RedmineProject
    For Each val In json("memberships")
      Set p = New RedmineProject
      p.id = Int(val("project")("id"))
      p.name = val("project")("name")
      pMemberships.Add p
    Next
  End If
  
  If json.Exists("groups") Then
    Dim g As RedmineGroup
    For Each val In json("groups")
      Set g = New RedmineGroup
      g.load val
      pGroups.Add g
    Next
  End If
End Function

Public Function generateJson( _
                  Optional ByVal password As String = "", _
                  Optional ByVal auth_source_id As String = "", _
                  Optional ByVal mail_notification As String = "", _
                  Optional ByVal mulst_change_password As Boolean = False, _
                  Optional ByVal generate_password As Boolean = False, _
                  Optional ByVal send_information As Boolean = False) As String
  generateJson = ""
  Dim root, user As Dictionary
  
  Set root = New Dictionary
  Set user = New Dictionary
  
  user.Add "login", Me.login
  user.Add "firstname", Me.firstname
  user.Add "lastname", Me.lastname
  user.Add "mail", Me.mail
  
  If password <> "" Then Call user.Add("password", password)
  If auth_source_id <> "" Then Call user.Add("auth_source_id", auth_source_id)
  If mail_notification <> "" Then Call user.Add("mail_notification", mail_notification)
  If mulst_change_password Then Call user.Add("mulst_change_password", mulst_change_password)
  If generate_password Then Call user.Add("generate_password", generate_password)
  
  root.Add "user", user
  If send_information Then root.Add "send_information", send_information
  
  generateJson = JsonConverter.ConvertToJson(root)
End Function

