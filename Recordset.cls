VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetOracleData"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mORADB_Conn     As ADODB.Connection
Private mORADB_RS       As ADODB.Recordset
Private mORADB_Cmd      As ADODB.Command
Private msUserID        As String
Private msPwd           As String
Private mbConnected     As Boolean
Private msErrNum        As String
Private msErrDesc       As String
Private msErrLoc        As String
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Purpose:         To serve as common object that uses a stored procedure and passes back the results
'*                  of the stored procedure.
'* Usage:           This module is used as follows:
'*
'*                  First, the userid and password must be passed to this object by using the SetUserIDPwd method.
'*                  In the calling object, a set of constants must be declared and then used to indicate the type
'*                  of parameter being passed.
'*                  The following constants must be declared in the calling object
'*
'*                  Private Const kRecordset    As Integer = 8
'*                  Private Const kString       As Integer = 129
'*                  Private Const kResultInOut  As Integer = 3
'*                  Private Const kInput        As Integer = 1
'*                  Private Const kOutput       As Integer = 2
'*
'*                  The calling object must generate an array of (x, 4) where x = the number of
'*                  parameters that are required by the stored procedure.
'*
'*                  The calling object must then set an ADODB. recordset = .GetSP_Data(stored procedure name, passed array)
'*
'*                  After using this method, the calling object must check the error code property of this object to see if it is 0.
'*                  In this case, the calling object can expect that the recordset object contains the value of the recordset returned
'*                  by the stored procedure and that the array contains any returned parameters.
'*                  The method for creating the array is as follows:
'*
'*                  Each parameter to be passed to the stored procedure should be created as a row of the array and in the same order
'*                  that the stored procedure expects to see it.  The following is an example
'*
'*                  lDataArray(1,1) = kInput       -----This indicates whether the parameter is input or output
'*                  lDataArray(1,2) = "xyz"        -----Variable Name
'*                  lDataArray(1,3) = kString      -----Variable Type, string or recordset
'*                  lDataArray(1,4) = 50           -----Variable Size
'*
'*                  The output variable will be similar and when the data array returns to the calling object, the (x,2) position
'*                  will contain the value of the variable returned by the stored procedure and the (x,1) position will contain the
'*                  variable name as was passed to this object.
'*
'* Description      This object uses looping to create command parameters for an ADO object and a SQL statement that contains the
'*                  name of the stored procedure and as many "?" marks as the upper bound of the array.  The values of the variable size,
'*                  type and whether input or output are used to create and append the parameters to the command object correctly.
'*                  This object first makes a connection and then keeps the connection open until the object is set to nothing.  Each time
'*                  a call is made to this object, it checks to see the state of the connection, except for the first time when it actually
'*                  makes the connection.
'*
'*                  NOTE:  always pass recordset as first parameter
'************************************************************************************************************************************************

Private Function SQLBuilder(ByVal SPName As String, ByRef SPDataArray() As Variant) As String
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo SQLBuilder_Err

    Dim intStringBuilder As Integer
    Dim lSQL As String

    'Create the beginning of the SQL statement using the stored procedure name
    lSQL = "{Call " & SPName & "(?"
    
    'Since the first ? is already inserted in SQL, loop one less than full upperbound
    For intStringBuilder = 1 To (UBound(SPDataArray, 1) - 1)
        lSQL = lSQL & ", ?"
    Next intStringBuilder
    
    'Add the closing parantheses
    lSQL = lSQL & ")}"
    
    'Set the return value
    SQLBuilder = lSQL
    
SQLBuilder_Exit:
    Exit Function
SQLBuilder_Err:
    msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
    msErrDesc = msErrDesc & Chr(13) & Err.Description
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.SQLBuilder"
    Resume SQLBuilder_Exit
    
End Function

Private Sub CloseConnection()
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo CloseConnection_Err
    
    'Close the global connection object
    mORADB_Conn.Close

CloseConnection_Exit:
    Exit Sub
CloseConnection_Err:
        msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
        msErrDesc = msErrDesc & Chr(13) & Err.Description
        msErrLoc = msErrLoc & Chr(13) & "Data Layer.CloseConnection"
        Resume CloseConnection_Exit
    
End Sub

Private Sub ParameterBuilder(ByRef SPDataArray() As Variant)
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo ParameterBuilder_Err

    Dim intParamBuild As Integer
    Dim paramString As String
    

    'Loop through the array and create/append the parameters
    For intParamBuild = 1 To UBound(SPDataArray, 1)
        'If third position indicates 8 which is the constant for recordset, create the parameter
        If SPDataArray(intParamBuild, 3) = 8 Then
            'Create and append the parameters in single statement
            mORADB_Cmd.Parameters.Append mORADB_Cmd.CreateParameter(SPDataArray(intParamBuild, 2), SPDataArray(intParamBuild, 3), SPDataArray(intParamBuild, 1))
        Else
            'If first position indicates 1, then it is an input
            If SPDataArray(intParamBuild, 1) = 1 Then
                'Create and append the parameters in single statement
                mORADB_Cmd.Parameters.Append mORADB_Cmd.CreateParameter("input", SPDataArray(intParamBuild, 3), SPDataArray(intParamBuild, 1), SPDataArray(intParamBuild, 4), SPDataArray(intParamBuild, 2))
            'If first position indicates 2, then it is an output
            ElseIf SPDataArray(intParamBuild, 1) = 2 Then
                'Create and append the parameters in single statement
                mORADB_Cmd.Parameters.Append mORADB_Cmd.CreateParameter(SPDataArray(intParamBuild, 2), SPDataArray(intParamBuild, 3), SPDataArray(intParamBuild, 1), SPDataArray(intParamBuild, 4), adEmpty)
            End If
        End If
    Next intParamBuild
    
ParameterBuilder_Exit:
    Exit Sub
ParameterBuilder_Err:
    msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
    msErrDesc = msErrDesc & Chr(13) & Err.Description
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.ParameterBuilder"
    Resume ParameterBuilder_Exit
    
End Sub

Private Sub RetParamBuilder(ByRef SPDataArray() As Variant)
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo RetParamBuilder_Err

    Dim intParamBuild As Integer

    'Build return array using the command parameters which now contain values returned by the recordset
    For intParamBuild = 1 To UBound(SPDataArray, 1)
        If SPDataArray(intParamBuild, 1) = 2 Then   'This is an output
            SPDataArray(intParamBuild, 1) = SPDataArray(intParamBuild, 2)   'So move the name into first position
            SPDataArray(intParamBuild, 2) = mORADB_Cmd.Parameters(SPDataArray(intParamBuild, 2)).Value  'and put the returned value in
                                                                                                        'second position
        End If
    Next intParamBuild

RetParamBuilder_Exit:
    Exit Sub
RetParamBuilder_Err:
    msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
    msErrDesc = msErrDesc & Chr(13) & Err.Description
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.RetParamBuilder"
    Resume RetParamBuilder_Exit

End Sub
Public Function GetSP_Data(ByVal SPName As String, ByRef SPDataArray() As Variant) As ADODB.Recordset
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo GetSP_Data_Err
    
    Dim lSQL As String
    
    'Main method
    
    Screen.MousePointer = vbHourglass
    
    'Set errors from last usage to zero
    msErrNum = 0
    msErrDesc = ""
    msErrLoc = ""

    'Check the connection and if OK then continue or else raise error here locally and report back to calling object
    
    If CheckConnection Then
        
        'Build the SQL statement
        lSQL = SQLBuilder(SPName, SPDataArray)
        
        'Call the procedure that sets the properties of the command object, including cleaning old appended parameters etc.
        Call CreateData(lSQL, SPDataArray)
        
        'Set the ADODB recordset = to the execute method of the command object
        Set GetSP_Data = mORADB_Cmd.Execute
        
        'Call the return builder to put values from the returned parameters into the array passed (byref)
        Call RetParamBuilder(SPDataArray)
    Else
        
        'Raise an error (Note that this is different than errors of stored procedure.  This error is if there is a problem
        'during connection
        Err.Raise -5001
    End If
    
    'The passed array now contains err_cd and err_txt which contains any errors from within the stored procedure itself.
    'Since the calling object will check for these error numbers, make sure to set them so that stored procedure errors
    'are also seen immediately by the calling object.
    msErrNum = CStr(SPDataArray((UBound(SPDataArray, 1) - 1), 2))
    msErrDesc = IIf(IsNull(SPDataArray(UBound(SPDataArray, 1), 2)), "", SPDataArray(UBound(SPDataArray, 1), 2))
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.GetSP_Data"
    
    
    If msErrDesc = "Err_txt" Then
        msErrDesc = "Problem when executing command object for ADO recordset"
    End If
    
GetSP_Data_Exit:
        Screen.MousePointer = vbDefault
        Exit Function
    
GetSP_Data_Err:

    'Somewhat of defunct code.......being cautious!!!!
    If Err.Number = -5001 Then
        msErrNum = CStr(-5001)
        msErrDesc = msErrDesc & Chr(13) & "Unable to make connection"
        msErrLoc = "Data Layer.GetSP_Data"
    Else
        msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
        msErrDesc = msErrDesc & Chr(13) & Err.Description
        msErrLoc = msErrLoc & Chr(13) & "Data Layer.GetSP_Data"
    End If
    Resume GetSP_Data_Exit

End Function

Private Sub Class_Initialize()
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

    'Set global flag explicitly when class is initialized
    mbConnected = False

End Sub

Private Sub Class_Terminate()
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

    'Explicitly close the connection
    Call CloseConnection

End Sub

Private Sub ParameterClear()
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo ParameterClear_Err

    Dim intParamClear As Integer
    
    'Clear out all the parameters of the command object
    For intParamClear = 1 To mORADB_Cmd.Parameters.Count
        mORADB_Cmd.Parameters.Delete (0)
    Next intParamClear
    
ParameterClear_Exit:
    Exit Sub
ParameterClear_Err:
    msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
    msErrDesc = msErrDesc & Chr(13) & Err.Description
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.ParameterClear"
    Resume ParameterClear_Exit
    

End Sub

Public Sub SetUserIDPwd(ByVal UserID As String, ByVal Pwd As String)
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

    msUserID = UserID
    msPwd = Pwd
    
End Sub

Private Function CheckConnection() As Boolean
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************
On Error GoTo CheckConnection_Err

    'Check global flag.  If true, that means, connection has once been opened
    If mbConnected Then
        If (mORADB_Conn.State = adStateOpen) Then
            'Check for connection state
            CheckConnection = True
            Exit Function
        End If
    End If
        
    'If global flag is false or even if true and connection has been broken, the rest of this code will execute
    
    'Set global objects
    Set mORADB_Conn = New ADODB.Connection
    Set mORADB_RS = New ADODB.Recordset
    Set mORADB_Cmd = New ADODB.Command
    
    'Open connection
    mORADB_Conn.ConnectionString = "DSN=Oracle_csd;UID=" & msUserID & ";PWD=" & msPwd
    mORADB_Conn.CursorLocation = adUseClient
    mORADB_Conn.Open
       
    'Set flags
    mbConnected = True
    CheckConnection = True
    
CheckConnection_Exit:
    Exit Function

CheckConnection_Err:
    msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
    msErrDesc = msErrDesc & Chr(13) & Err.Description
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.CheckConnection"
    mbConnected = False
    CheckConnection = False
    Resume CheckConnection_Exit

End Function

Private Sub CreateData(ByVal lSQL As String, ByRef SPDataArray())
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************

On Error GoTo CreateData_Err

    'First clear out existing parameters
    Call ParameterClear
    
    'Then build new parameters
    Call ParameterBuilder(SPDataArray)
    
    'Set command objects connection equal to the live connection
    mORADB_Cmd.ActiveConnection = mORADB_Conn
    
    'Set the command objects text = SQL built earlier
    mORADB_Cmd.CommandText = lSQL
    
    'Set the comand objects type to text
    mORADB_Cmd.CommandType = adCmdText

CreateData_Exit:
    Exit Sub

CreateData_Err:
    msErrNum = msErrNum & Chr(13) & CStr(Err.Number)
    msErrDesc = msErrDesc & Chr(13) & Err.Description
    msErrLoc = msErrLoc & Chr(13) & "Data Layer.CreateData"
    Resume CreateData_Exit

End Sub
Public Property Get ErrorCode() As Variant
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************
    
    ErrorCode = msErrNum
End Property

Public Property Get ErrorDescription() As Variant
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************
    
    ErrorDescription = msErrDesc
End Property
Public Property Get ErrorLocation() As Variant
'************************************************************************************************************************************************
'* Original Author: Aparajit Agarwal
'* Created:         4/1999
'************************************************************************************************************************************************
    
    ErrorLocation = msErrLoc
End Property

