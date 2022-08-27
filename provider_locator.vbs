Option Explicit

ExecuteGlobal read("./VbsJson.vbs")
ExecuteGlobal read("./urlencode.vbs")

Dim jsonParser : Set jsonParser = New VbsJson
Dim objShell   : Set objShell   = CreateObject("WScript.Shell")
Dim shell      : Set shell      = CreateObject("WScript.Shell")
Dim wowUpData  :     wowUpData  = shell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""").StdOut.ReadLine ' https://stackoverflow.com/a/21565999

Dim addons, addon, text

text = read(wowUpData)
addons = jsonParser.Decode(text).Items

For Each addon In addons
  If addon("providerName") = "Unknown" then
    objShell.run("http://www.google.com/search?q=" & URLEncode(addon("name") & " wow addon"))
  End If
Next

' https://stackoverflow.com/a/10615831
Function read(fileName)
  Dim fileSystem, file

  Set fileSystem = WScript.CreateObject("Scripting.Filesystemobject")
  Set file = fileSystem.OpenTextFile(fileName)

  read = file.ReadAll
  file.Close
End Function
