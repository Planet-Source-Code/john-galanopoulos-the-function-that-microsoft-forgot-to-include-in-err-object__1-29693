Attribute VB_Name = "modApiErr"
Option Explicit
'John Galanopoulos,  GreekThought@yahoo.gr

'This source returns an error description from the system
'You pass the Err.LastDLLError to the function
'LastDLLErrorDescription after an API call.
'This is a great debugger for your API function-returned-troubles and i really
'can't understand why Microsoft didn't included in it's Err Object

'..but make no mistake :
'Err.Number       -> Err.Description
'Err.LastDLLError -> LastDLLErrorDescription()

'Any comments or suggestions are accepted and much appreciated.
'(I will repeat myself but(!) i see from my other projects that 700 ppl downloaded the source
'and only 2 or 3 made a comment or even voted.)


'Although this example looks harmless, you never know with windows
'so use at your own risk.

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
'Specifies that the function should search the system message-table
'resource(s) for the requested message

Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
'Specifies that insert sequences in the message definition
'are to be ignored and passed through to the output buffer unchanged.
'This flag is useful for fetching a message for later formatting.
'If this flag is set, the Arguments parameter is ignored.

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'This API Function returns the number of characters stored in the output buffer,
'excluding the terminating null character, and this indicates success.
'if it returns Zero, this indicates execution failure.

Public Function LastDLLErrorDescription(LastDLLError As Long) As String
  
   Dim strBuffer   As String
   Dim LenBuff     As Long
   
   Dim cRes        As Long
   
   LenBuff = 256
   strBuffer = Space$(LenBuff)
   
   
   cRes = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, LastDLLError, 0&, strBuffer, LenBuff, 0&)
   'Now here is the mistake that many do. The FormatMessage function, if succesfull, returns
   'the buffer length of the strBuffer  . No need to use special routines to start cleaning
   'the buffer from trash or vbcrlfs.
   
   'Look how easy this works
      
   If cRes = 0 Then
      strBuffer = "FormatMessage API execution Error. Couldn't fetch error description."
     Else
      strBuffer = Left$(strBuffer, cRes - 2) '(Althought Microsoft says that it excludes the
                                             'vbcrlfs, in my test form it returned 2 vbcrlfs)
   End If
   
   
   LastDLLErrorDescription = strBuffer

End Function

