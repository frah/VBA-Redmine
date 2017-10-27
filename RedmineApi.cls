VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Redmineエンドポイント
Public BaseUri As String
'Redmine API Key
Public ApiKey As String

Private json As Dictionary

'Constructor
Private Sub Class_Initialize()
  BaseUri = vbNullString
  ApiKey = vbNullString
End Sub
'Destructor
Private Sub Class_Terminate()
  Set json = Nothing
End Sub

Public Function CreateUser( _
                  ByRef user As RedmineUser, _
                  Optional ByVal password As String = "", _
                  Optional ByVal auth_source_id As String = "", _
                  Optional ByVal mail_notification As String = "", _
                  Optional ByVal mulst_change_password As Boolean = False, _
                  Optional ByVal generate_password As Boolean = False, _
                  Optional ByVal send_information As Boolean = False) As Boolean
  Dim json As String
  json = user.generateJson(password, auth_source_id, mail_notification, must_change_password, generate_password, send_information)
  CreateUser = SendApiRequest("POST", "users", requestBody:=json)
End Function

'ユーザ情報取得
Public Function GetUsers( _
                  Optional ByVal userId As Integer = -1, _
                  Optional ByVal incMemberships As Boolean = True, _
                  Optional ByVal incGroups As Boolean = True, _
                  Optional ByVal status As Integer = 1, _
                  Optional ByVal name As String = "", _
                  Optional ByVal group_id As Integer = -1) As Collection
  Dim reqPath, incStr As String
  Dim params As Object
  
  reqPath = "users"
  If userId >= 0 Then
    reqPath = reqPath & "/" & userId
  End If
  
  incStr = ""
  If incMemberships Then incStr = incStr & ",memberships"
  If incGroups Then incStr = incStr & ",groups"
  
  Set params = CreateObject("Scripting.Dictionary")
  If incStr <> "" Then params.Add "include", Mid(incStr, 2)
  If status <> 1 Then params.Add "status", status
  If name <> "" Then params.Add "name", name
  If group_id <> -1 Then params.Add "group_id", group_id
  
  If SendApiRequest("GET", reqPath, requestParam:=params) Then
    Dim ret As Collection
    Dim r As RedmineUser
    Set ret = New Collection
    
    If json.Exists("user") Then
      'Single result
      Set r = Nothing
      Set r = New RedmineUser
      r.load json("user")
      ret.Add r
    Else
      'Multiple result
      Dim u As Object
      For Each u In json("users")
        Set r = Nothing
        Set r = New RedmineUser
        r.load u
        ret.Add r
      Next
    End If
    Set GetUsers = ret
  Else
    Set GetUsers = Nothing
  End If
End Function

Public Function GetGroups( _
                  Optional ByVal groupId As Integer = -1, _
                  Optional ByVal incUsers As Boolean = False, _
                  Optional ByVal incMemberships As Boolean = False) As Collection
  Dim reqPath, incStr As String
  Dim params As Object
  
  reqPath = "groups"
  If groupId >= 0 Then
    reqPath = reqPath & "/" & groupId
  End If
  
  incStr = ""
  If incMemberships Then incStr = incStr & ",memberships"
  If incUsers Then incStr = incStr & ",users"
  
  Set params = CreateObject("Scripting.Dictionary")
  If incStr <> "" Then params.Add "include", Mid(incStr, 2)
  
  If SendApiRequest("GET", reqPath, requestParam:=params) Then
    Dim ret As Collection
    Dim r As RedmineUser
    Set ret = New Collection
    
    If json.Exists("group") Then
      'Single result
      Set r = Nothing
      Set r = New RedmineGroup
      r.load json("group")
      ret.Add r
    Else
      'Multiple result
      Dim u As Object
      For Each u In json("groups")
        Set r = Nothing
        Set r = New RedmineGroup
        r.load u
        ret.Add r
      Next
    End If
    Set GetGroups = ret
  Else
    Set GetGroups = Nothing
  End If
End Function

'プロジェクト情報取得
' @param projectId: プロジェクトID
' @param incTrackers: プロジェクトのトラッカー一覧も返す
' @param incCategories: プロジェクトのカテゴリ一覧も返す
' @param incModules: プロジェクトのモジュール一覧も返す
' @param incActivities: プロジェクトの直近アクティビティも返す
' @return: 結果のXMLツリー
Public Function GetProjects( _
                  Optional ByVal projectId As String = "", _
                  Optional ByVal incTrackers As Boolean = False, _
                  Optional ByVal incCategories As Boolean = False, _
                  Optional ByVal incModules As Boolean = False, _
                  Optional ByVal incActivities As Boolean = False) As Collection
  Dim reqPath, incStr As String
  Dim params As Object
  
  reqPath = "projects"
  If projectId <> "" Then
    reqPath = reqPath & "/" & projectId
  End If
  
  If incTrackers Then incStr = incStr & ",trackers"
  If incCategories Then incStr = incStr & ",issue_categories"
  If incModules Then incStr = incStr & ",enabled_modules"
  If incActivities Then incStr = incStr & ",time_entry_activities"
  
  If incStr <> "" Then
    Set params = Nothing
  Else
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "include", Mid(incStr, 2)
  End If
  
  If SendApiRequest("GET", reqPath, requestParam:=params) Then
    Dim ret As Collection
    Dim r As RedmineProject
    Set ret = New Collection
    
    If json.Exists("project") Then
      'Single result
      Set r = Nothing
      Set r = New RedmineProject
      r.load json("project")
      ret.Add r
    Else
      'Multiple result
      Dim u As Object
      For Each u In json("projects")
        Set r = Nothing
        Set r = New RedmineProject
        r.load u
        ret.Add r
      Next
    End If
    Set GetProjects = ret
  Else
    Set GetProjects = Nothing
  End If
End Function

'RedmineへAPIリクエスト送信
' @param method: メソッド(GET,POST,PUT,DELETE)
' @param requestPath: リクエスト先のリソース
' @param requestBody: リクエストボディ
' @return: リクエストの成否
Public Function SendApiRequest( _
                  ByVal method As String, _
                  ByVal requestPath As String, _
                  Optional ByVal requestParam As Object = Nothing, _
                  Optional ByVal requestBody As String = "") As Boolean
  Dim key As Variant
  Dim postUri As String
  Dim req As Object
  
  SendApiRequest = False
  
  If BaseUri = vbNullString Or ApiKey = vbNullString Then
    '必須パラメータなし
    Exit Function
  End If
  
  'HttpRequestオブジェクト作成
  Set req = CreateHttpRequest()
  
  'URL組み立て
  postUri = Me.BaseUri & requestPath & ".json"
  'パラメタ付与
  If Not requestParam Is Nothing And TypeName(requestParam) = "Dictionary" Then
    If LBound(requestParam.Keys) >= 0 Then
      postUri = postUri & "?"
      For Each key In requestParam
        postUri = postUri & key & "=" & UrlEncodeUtf8(requestParam.Item(key)) & "&"
      Next
    End If
  End If

SkipBuildPostData:
  'リクエスト送信
  req.Open method, postUri, False
  req.setRequestHeader "Content-Type", "application/json"
  req.setRequestHeader "X-Redmine-API-Key", Me.ApiKey
  req.send requestBody
  
  If Not req.status < 300 Then
    Debug.Print "StatusCode: " & req.status
    Debug.Print "Response: " & req.responseText
    Exit Function
  End If
  
  'レスポンスパース
  Set json = Nothing
  Set json = JsonConverter.ParseJson(req.responseText)
  
  SendApiRequest = True
End Function

'URLエンコード文字列を返す
Private Function UrlEncodeUtf8(ByVal str As String) As String
  Dim script As Object
  Set script = CreateObject("ScriptControl")
  script.Language = "Jscript"
  UrlEncodeUtf8 = script.CodeObject.encodeURIComponent(str)
End Function

'テキストの文字コード変換
' @param text: 変換対象のバイト列
' @param fromCharset: 変換前の文字コード
' @param toCharset: 変換先の文字コード（デフォルト unicode）
' @return: 変換後の文字列
Private Function ConvTextEncoding( _
                    ByVal text, _
                    ByVal fromCharset As String, _
                    Optional ByVal toCharset As String = "unicode") As String
  Dim convText As String
  With CreateObject("ADODB.Stream")
    .Open
    .Type = adTypeText
    .Charset = toCharset
    .WriteText text
    .Position = 0
    .Type = adTypeText
    .Charset = fromCharset
    
    On Error GoTo myLabel
    convText = .ReadText()
    ConvTextEncoding = Mid(convText, 3, Len(convText))
    
    On Error GoTo 0
  End With
  Exit Function

myLabel:
  ConvTextEncoding = StrConv(text, vbUnicode, 1041)
End Function

Private Function SaveAsFile(ByVal datas, ByVal path As String) As Boolean
  With CreateObject("ADODB.Stream")
    .Type = adTypeBinary
    .Open
    .Write datas
    .SaveToFile path, adSaveCreateOverWrite
    .Close
  End With
End Function

Private Function CreateHttpRequest() As Object
'WinHttpRequest/XMLHTTPRequestオブジェクト作成
'http://www.f3.dion.ne.jp/~element/msaccess/AcTipsWinHTTP1.html 参考
  Dim progIDs As Variant
  Dim ret As Object
  Dim i As Long
   
  Set ret = Nothing '初期化
  progIDs = Array("WinHttp.WinHttpRequest.5.1", _
                  "WinHttp.WinHttpRequest.5", _
                  "WinHttp.WinHttpRequest", _
                  "Msxml2.ServerXMLHTTP.6.0", _
                  "Msxml2.ServerXMLHTTP.5.0", _
                  "Msxml2.ServerXMLHTTP.4.0", _
                  "Msxml2.ServerXMLHTTP.3.0", _
                  "Msxml2.ServerXMLHTTP", _
                  "Microsoft.ServerXMLHTTP", _
                  "Msxml2.XMLHTTP.6.0", _
                  "Msxml2.XMLHTTP.5.0", _
                  "Msxml2.XMLHTTP.4.0", _
                  "Msxml2.XMLHTTP.3.0", _
                  "Msxml2.XMLHTTP", _
                  "Microsoft.XMLHTTP")
  On Error Resume Next
  For i = LBound(progIDs) To UBound(progIDs)
    Set ret = CreateObject(progIDs(i))
    If Not ret Is Nothing Then Exit For
  Next
  On Error GoTo 0
  Set CreateHttpRequest = ret
End Function

