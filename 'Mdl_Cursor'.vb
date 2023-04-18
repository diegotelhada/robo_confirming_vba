'Mdl_Cursor'

Option Compare Database

Option Explicit

Public Const IDC_APPSTARTING = 32650&
Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZE = 32640&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&

Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" _
  (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Declare Function LoadCursorFromFile Lib "user32" Alias _
  "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Declare Function SetCursor Lib "user32" _
  (ByVal hCursor As Long) As Long

Const curNAME = "harrow.cur"
Private mhCursor As Long
Private mstrCursorPath As String
Private Const ERR_INVALID_CURSOR = vbObjectError + 3333

Function MouseCursor(CursorType As Long)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function

Function PointM(strPathToCursor As String)
    If mhCursor = 0 Then
        mhCursor = LoadCursorFromFile(strPathToCursor)
    End If
    Call SetCursor(mhCursor)
End Function

Public Sub GetCursor()

On Error GoTo ErrHandler
    If Len(mstrCursorPath) = 0 Then
      mstrCursorPath = CurrentDb.Name
       mstrCursorPath = Left(mstrCursorPath, InStr(mstrCursorPath, Dir(mstrCursorPath)) - 1)
        mstrCursorPath = mstrCursorPath & curNAME
        If Len(Dir(mstrCursorPath)) = 0 Then
            mstrCursorPath = vbNullString
        End If
   End If
    If Len(mstrCursorPath) = 0 Then
        Err.Raise ERR_INVALID_CURSOR
   Else
        PointM (mstrCursorPath)
    End If
ExitHere:
    Exit Sub
ErrHandler:
   With Err
      If .Number = ERR_INVALID_CURSOR Then
        MsgBox "Error: " & .Number & vbCrLf & _
           "Invalid Cursor type", _
                vbCritical Or vbOKOnly, _
                "Cursor Function"
        Else
          MsgBox "Error: " & .Number & vbCrLf & _
                   .Description, _
                vbCritical Or vbOKOnly, _
                "Cursor Function"
        End If
    End With
   Resume ExitHere
End Sub







