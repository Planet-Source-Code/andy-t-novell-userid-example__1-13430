Attribute VB_Name = "Module1"
' Following declaration was taken from
' Novell Libraries for Visual Basic
' available at http://developer.novell.com/ndk/download.htm

' Useful Novell Visual Basic Links:
'   http://developer.novell.com/ndk/vblib.htm - Novell Libraries for Visual Basic
'   http://developer.novell.com/support/sample/areas/vbs.htm - Novell Visual Basic Code Examples

' The Novell developer website has a lot of really good VB code examples.
' They also have a newsgroup where the author of most of the code answers
' questions and gives examples

Public Const DCK_TREE_NAME = 11

Public Type Tree_Name_T
    tname(64) As Byte
End Type

Declare Function NWDSCreateContextHandle Lib "NETWIN32" _
    (context As Long) As Long

Declare Function NWDSScanConnsForTrees Lib "NETWIN32" _
    (ByVal context As Long, ByVal numOfPtrs As Long, _
     numOfTrees As Long, treeBufPtrs As Long) As Long

Declare Function NWDSSetContext Lib "NETWIN32" _
    (ByVal context As Long, ByVal key As Long, _
     ByVal value As Long) As Long

Declare Function NWDSWhoAmI Lib "NETWIN32" _
    (ByVal context As Long, ByVal objectName As Long) As Long
