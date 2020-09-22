VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Get Novell Data"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Data"
      Height          =   375
      Left            =   1313
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Novell Tree Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Novell Login Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' I cannot take credit for this code
' Most of this code is taken from examples posted on the Novell Developer Website
' See Module1 for Novell Visual Basic Links
' Just passing the information on, that's all

Dim mstrTreeName As String

Private Sub Command1_Click()
    Text1.Text = sGetUserID
    Text2.Text = mstrTreeName
End Sub

Private Function sGetUserID() As String

    Dim retCode As Long
    Dim byteName(127) As Byte
    Dim name As String
    Dim i As Integer
    Dim contextHandle As Long, treePointers(15) As Long
    Dim treeNames(15) As Tree_Name_T
    Dim numOfTrees As Long

' Part I. - getting connected DS tree names
    retCode = NWDSCreateContextHandle(contextHandle)
    If retCode <> 0 Then
        MsgBox "NWDSCreateContextHandle failed, E=" + retCode, vbCritical
    Else
' I.a - We need to initialize structure for authenticated tree names
'       Let`s say there won`t be more then 16 trees...
        For i = 0 To 15
            treePointers(i) = VarPtr(treeNames(i).tname(0))
        Next i
' I.b - Following function searches for connected DS tree names
        retCode = NWDSScanConnsForTrees(contextHandle, 16, numOfTrees, treePointers(0))
        If retCode <> 0 Then
            MsgBox "NWDSScanConnsForTrees failed, E=" + retCode, vbCritical
        Else
            For i = 0 To numOfTrees - 1
                Call ByteArrayToString(treeNames(i).tname, name)
                mstrTreeName = name
' I.c - Now we have DS tree name and need to know our user name
'        in this DS tree
                Call StringToByteArray(name + Chr(0), byteName)
                retCode = NWDSSetContext(contextHandle, DCK_TREE_NAME, VarPtr(byteName(0)))
                retCode = NWDSWhoAmI(contextHandle, VarPtr(byteName(0)))
                Call ByteArrayToString(byteName, name)
                ' This just takes the 'CN=' out of the user ID
                sGetUserID = UCase(Mid(name, 4))
            Next i
        End If
    End If
End Function

Private Sub ByteArrayToString(src() As Byte, dest As String)
Dim i As Integer
    i = 0
    dest = ""
    While src(i) <> 0
        dest = dest + Chr(src(i))
        i = i + 1
    Wend
End Sub

Private Sub StringToByteArray(src As String, dest() As Byte)

' Following For-Next loop  should run to 0x0 char only
'   but we do not care if it runs longer
    Dim i As Integer

    For i = 0 To Len(src) - 1
        dest(i) = CByte(Asc(Mid(src, i + 1, 1)))
    Next i
End Sub
