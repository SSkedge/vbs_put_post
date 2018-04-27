'@Author Steven Skedge
'Some content researched from external sources 
'Date July 2016
'PUT
'A class to encode files and parse to JSON
'Made for the command line
'
Option Explicit

Const TypeBinary = 1
Const ForReading = 1, ForWriting = 2, ForAppending = 8
  
Dim arguments, inFile, outFile, fileNameArray

' Gets file from passed in args
Set arguments = WScript.Arguments

'in cmd putfile.vbs nameoffile
inFile = arguments(0)

'For decoded file
'outFile = "new_" & inFile


Dim inByteArray, base64Encoded, base64Decoded, outByteArray,i,encodeArray()
ReDim Preserve encodeArray (arguments.Count)

For i = 0 to arguments.count - 1
inFile = arguments(i) 
 
inByteArray = readBytes(inFile)
base64Encoded = encodeBase64(inByteArray)
encodeArray(i) = base64Encoded

Next

PostJSON(encodeArray)
  
'Wscript.echo "Base64 encoded: " + base64Encoded

'-------------------------------------------
'Uncomment for decoding to file
'base64Decoded = decodeBase64(base64Encoded)
'writeBytes outFile, base64Decoded
'-------------------------------------------
 
'
'Reads the bytes of the file
'
'
private function readBytes(file)
  dim inStream
  ' ADODB stream object used
  set inStream = WScript.CreateObject("ADODB.Stream")
  ' open with no arguments makes the stream an empty container 
  inStream.Open
  inStream.type= TypeBinary
  inStream.LoadFromFile(file)
  readBytes = inStream.Read()
end function
  
'
'Encodes to base 64
'
'
private function encodeBase64(bytes)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  encodeBase64 = EL.Text
end function


'-------------------------------------------
'Uncomment for decoding to file 
'
'
'private function decodeBase64(base64)
 ' dim DM, EL
  'Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  'Set EL = DM.createElement("tmp")
  'EL.DataType = "bin.base64"
  ' Set encoded String, get bytes
  'EL.Text = base64
  'decodeBase64 = EL.NodeTypedValue
'end function
'-------------------------------------------
  
'
'Writes bytes 
'
'
private Sub writeBytes(file, bytes)
  Dim binaryStream
  Set binaryStream = CreateObject("ADODB.Stream")
  binaryStream.Type = TypeBinary
  'Open the stream and write binary data
  binaryStream.Open
  binaryStream.Write bytes
  'Save binary data to disk
  binaryStream.SaveToFile file, ForWriting
End Sub

'
'Function to postJSON
'
'
Public Function PostJSON (encodeArray)
	
  Dim objHTTP, URL, objXmlHttpMain, json, fso, stdout,stderr,strJson,j,strSendJSON, strFinal,fileName
  
  
  Set fso = CreateObject ("Scripting.FileSystemObject") 
  Set stdout = fso.GetStandardStream (1) 
  Set stderr = fso.GetStandardStream (2) 

  Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
  
  'Url specific to the server PUT places in envelope and must be an existing envelope ID
  URL = "https://jsonplaceholder.typicode.com/puts"
  objHTTP.Open "POST", URL, False
  
  'Not all accurate header content, Script must match this shape
  objHTTP.setRequestHeader "Authorization", "Bearer 4321"
  objHTTP.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
  objHTTP.setRequestHeader "ApiKey", "1234"
  objHTTP.setRequestHeader "UserToken", "Bearer 4321"
  
  For j = 0 to UBound(encodeArray) - 1
  
  'Collects filename from arguments
  fileName = Wscript.arguments(j)
  strJson = "{" & Qu("Name") & ": " & Qu(fileName) & ", " &_
           Qu("File") & ": " & Qu(encodeArray(j)) & "}"
  If j = UBound(encodeArray) - 2 Then
  strSendJSON = strSendJSON + strJson + ","
  Else
  strSendJSON = strSendJSON + strJson
  End If
  
  Next
  
  'Adds square brackets for correct array shape
  strFinal = "[" + strSendJSON + "]"
  
  'prints the final body layout
  'Wscript.echo strFinal

  ' Send the json in correct format
  json = strFinal
  objHTTP.send (json)

  ' Output error message to std-error and happy message to std-out. Should
  ' simplify error checking
  If objHTTP.Status >= 400 And objHTTP.Status <= 599 Then
    WScript.echo "Error Occurred : " & objHTTP.status & " - " & objHTTP.statusText
      PostJSON = false
  Else
    WScript.echo "Success : " & objHTTP.status & " - " & objHTTP.ResponseText
      PostJSON = true
  End If
End Function

'
'Function to improve readability
'
Public Function Qu(ByVal s)
  Qu = Null
  If (VarType(s) = vbString) Then
    Qu = Chr(34) & CStr(s) & Chr(34)
  End If
End Function

    
