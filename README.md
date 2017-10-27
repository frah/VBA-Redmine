# VBA-Redmine

Redmine REST API library for VBA

## Examples

```vb
Dim redmine As RedmineApi
Dim users As Collection
Dim user As RedmineUser

Set redmine = New RedmineApi
redmine.BaseUri = "http://example.com/"
redmine.ApiKey = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

Set users = redmine.GetUsers()
For Each user In users
  Debug.Print user.login & ":" & user.firstname & user.lastname
Next
```

## Setup

1. Import [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) into your project (File > Import File)
2. Import all of *.cls and *.bas files into your project
3. Add `Dictionary` reference/class
   - Include a reference to "Microsoft Scripting Rumtime"
