VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtLastName 
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtTitle 
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtFirstName 
      Height          =   495
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCloseConnection 
      Caption         =   "Close Connection"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenConnection 
      Caption         =   "Open Connection"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblEOF 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblBOF 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "EOF"
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "BOF"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Title"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblConnectionString 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Connection String:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is an example of Connecting to a mySql Databse DSNless
'Most of the code is not mine.  I have just modified it to
'Show how to connect to mySql  it was from a
'very good database tutorial found on this site.


Option Explicit



Private WithEvents adoDataConn As ADODB.Connection
Attribute adoDataConn.VB_VarHelpID = -1
'connConnection is the connection with the DB.
'A more descriptive name is usually desired.  For example:
'If you are connecting with the Northwind DB, connNWind
'would be better.

'The WithEvents keyword means that you’ll be able code
'events for the connection and the record set.  Also, you
'will be able to find the declared object in a code
'window’s drop-down list and that each object will further
'provide its event procedures in the rightside drop-down
'of the code window.

Private WithEvents rsRecordSet As ADODB.Recordset
Attribute rsRecordSet.VB_VarHelpID = -1
'rsRecordSet is the recordset that will be used with the
'connection.A more descriptive name is usually desired.
'If you are connecting with the employees table,
'rsEmployees would be better.

Dim mblnAddMode As Boolean
'Used to determine whether data should be displayed
'or added.

'***********
'**  Create cmdOpenConnection &
'***********  cmdCloseConnection.

Private Sub cmdOpenConnection_Click()
    'Remember, these steps can be in another sub such
    'as the Form_Load event.
        

    
    Dim strConnect As String
    'This is your connection string.  It will contain
    'information about the provider and the path to
    'the database.
    
    Dim strProvider As String
    'In order to keep from typing long strings, I am
    'breaking the connection string into smaller parts.
    'It should be easier to read this way.
    
    Dim strDataSource As String
    'See note for strProvider.
    
    Dim strDataBaseName As String
    
    
    Dim usr_id As String ' the user id for the database
    Dim pass As String ' the password if used in your database
    Dim mySqlIP As String 'the ip address of the machine with the mySql
    mySqlIP = "127.0.0.1"  ' this is for localhost

    usr_id = "myID" ' user id
    pass = "myPass" 'password
    
    
    ' This is your connection string
    strConnect = "driver={MySQL};server=" & mySqlIP & ";uid=" & usr_id & ";pwd=" & pass & ";database=webcalendar"
    'The connection string is now made.
    
    Set adoDataConn = New ADODB.Connection
    'Preparing the connection object.
    
    
    
    adoDataConn.CursorLocation = adUseClient
    'Use a client side cursor because the data you will
    'be accessing will be on the client machine instead
    'of a server.
    
    adoDataConn.Open strConnect
    'Open the connection object.
    
    lblConnectionString.Caption = strConnect
    'Just in case you are interested in seeing the string.
    


    Set rsRecordSet = New ADODB.Recordset
    'Prepare the recordset.

    rsRecordSet.CursorType = adOpenStatic
    'The only type of curor that you can use with
    'a client side cursor location is adOpenStatic.
    
    rsRecordSet.CursorLocation = adUseClient
    'This application is using a client side cursor.
    
    rsRecordSet.LockType = adLockPessimistic
    'This guarantees that a record that is being edited
    'can be saved.
    
    rsRecordSet.Source = "Select * From tAdmin"  ' change to your table
    'Source should be a SQL statement indicating where to
    'retreive the data from.
    
    rsRecordSet.ActiveConnection = adoDataConn
    'The record set needs to know what connection to use.
    
    rsRecordSet.Open
    'Open the record set.  Opening the recordset will
    'cause the MoveComplete event for the recordset to fire.
    
    cmdOpenConnection.Enabled = False
    'Since the connection is now opened, one should not be
    'allowed to try to open it again.
    
    cmdCloseConnection.Enabled = True
    'Since the connection is open, one should be allowed
    'to close it
    
    'Now, the connection should be open and the recordset
    'ready to work with.

    lblBOF.Caption = rsRecordSet.BOF
    lblEOF.Caption = rsRecordSet.EOF
End Sub

'***********
'* *  Create txtFirstName, txtLastName, txtTitle
'***********  Set the locked property to true for each one.

Private Sub ClearControls()

'***********
'* *
'***********

    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtTitle.Text = ""
End Sub

Private Sub cmdCloseConnection_Click()
    
'***********
'* *
'***********
    
    adoDataConn.Close
    Set adoDataConn = Nothing
    
    cmdCloseConnection.Enabled = False
    'The user should not be allowed to close a connection
    'that is not open.
    
    cmdOpenConnection.Enabled = True
    'Since the connection is closed, it is okay to open
    'it again.

    Call ClearControls
    'Clear the textboxes.
    
    lblConnectionString.Caption = ""
    lblBOF.Caption = ""
    lblEOF.Caption = ""
End Sub

Private Sub LoadDataInControls()

'***********
'**
'***********

    If rsRecordSet.BOF = True Or rsRecordSet.EOF = True Then
        Exit Sub
        'If the pointer is at the end of the recordset of
        'at the before the first record, exit the sub and
        'show no data.
    End If
    
    'There are several methods of refering to field
    'contents. Here are examples of some different ways.
    
    txtFirstName.Text = rsRecordSet.Fields("Login").Value & " "
    'set these to fields in your database
    'Notice the " " appended to the end of the field.
    'This is necessary in order to prevent errors
    'loading if the field is empty.  At least a space will
    'be loaded into the text box.
    
    txtLastName.Text = rsRecordSet("Password1").Value & " "
    'The fields property is missing here.
    
    txtTitle.Text = rsRecordSet!Level & " "
    'This is the bang method.  Notice that there are no
    'quotes around the field name.
End Sub

'***********
'**
'***********
Private Sub rsRecordSet_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If mblnAddMode = False Then
        'If we are not in the add mode, load data in the
        'controls.
        
        Call LoadDataInControls
        'This event was fired by opening the recordset in
        'step six.
        
    End If
End Sub


'***********
'**
'***********

Private Sub cmdFirst_Click()

    If rsRecordSet.BOF = False Then
        rsRecordSet.MoveFirst
        'Move to the first record in the record set.
    ElseIf rsRecordSet.BOF = True _
        And rsRecordSet.EOF = True Then
        
        MsgBox "There is no data in the record set!", , "Oops"
    End If
    
    lblBOF.Caption = rsRecordSet.BOF
    lblEOF.Caption = rsRecordSet.EOF
End Sub

Private Sub cmdLast_Click()

    If rsRecordSet.EOF = False Then
        rsRecordSet.MoveLast
        'Move to the last record in the record set.
    ElseIf rsRecordSet.BOF = True _
        And rsRecordSet.EOF = True Then
        
        MsgBox "There is no data in the record set!", , "Oops"
    End If
    
    lblBOF.Caption = rsRecordSet.BOF
    lblEOF.Caption = rsRecordSet.EOF
End Sub

Private Sub cmdPrevious_Click()

    If rsRecordSet.BOF = False Then
        rsRecordSet.MovePrevious
        'Check to see if you are at the front of the record set.
        'If you are not, then you can move forward.
        
        If rsRecordSet.BOF = True Then
            'This will prevent the user from moving to
            'the BOF marker if he or she is on the first record.
            rsRecordSet.MoveFirst
        End If
    Else
        If rsRecordSet.EOF Then
            'Check to see if there is any data in the record set.
            MsgBox "There is no data in the record set!", , "Oops"
        Else
            rsRecordSet.MoveFirst
            'If the user is at the BOF marker, then move
            'to the first record.  There are several other
            'ways to handle this. For example, you could
            'loop the user to the last record by using
            'rsRecordset.MoveLast after the else statement.
            
        End If
    End If
    
    lblBOF.Caption = rsRecordSet.BOF
    lblEOF.Caption = rsRecordSet.EOF
End Sub

Private Sub cmdNext_Click()
    
    If rsRecordSet.EOF = False Then
        rsRecordSet.MoveNext
        
        If rsRecordSet.EOF Then
            'This will prevent the user from moving to
            'the EOF Marker if he or she is on the last
            'record.
            
            rsRecordSet.MoveLast
        End If
    Else
        If rsRecordSet.BOF Then
            'Check to see if there is any data in the recordset.
        
            MsgBox "There is no data in the record set!", , "Oops"
        Else
            rsRecordSet.MoveLast
            'Move the user to the last record if he or she
            'tries to move past the last record.  You
            'could let the user add a record by moving
            'past the EOF marker or you could loop back
            'to the first record.
        End If
    End If

    lblBOF.Caption = rsRecordSet.BOF
    lblEOF.Caption = rsRecordSet.EOF
End Sub

'***********
'*14th Step*    Add cmdAdd and cmdSave to the form.
'***********

'***********
'*15th Step*
'***********

Private Sub DisableNavigation()
    
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
End Sub

Private Sub EnableNavigation()
    
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
End Sub

'***********
'*16th Step*
'***********

Private Sub cmdAdd_Click()
    
    If cmdAdd.Caption = "&Add" And _
        cmdCloseConnection.Enabled = True Then
        'Allow adds when the connection is opened.
        
        cmdAdd.Caption = "&Cancel"
        'Change the caption to Cancel and use this button
        'to prevent a new record from being saved.
    
        cmdSave.Enabled = True
        'Since a new record is being added, the user
        'should be allowed to save the data.
        
        Call DisableNavigation
        'the user should not be allowed to navigate during an
        'add.
        
        mblnAddMode = True
        'We are now in addmode.
        
        Call ClearControls
        'This sub just clears the controls in preparation for
        'adding new data.
        
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
        'Coded later to add more stability.
        
        txtFirstName.Locked = False
        txtLastName.Locked = False
        txtTitle.Locked = False
        'Allow the user to enter items into the text boxes.
        
        txtFirstName.SetFocus
        'Go to the first field.
        
    ElseIf cmdAdd.Caption = "&Cancel" Then
        cmdAdd.Caption = "&Add"
        'If a user cancels, allow another add.
        
        cmdSave.Enabled = False
        'There will be no record to save.
        
        Call EnableNavigation
        'The user should have freedom to move now.
        
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        'Added later to provide more security.
        
        mblnAddMode = False
        'We are no longer in addmode.
        
        txtFirstName.Locked = True
        txtLastName.Locked = True
        txtTitle.Locked = True
        'Allow the user to enter items into the text boxes.
        
        If cmdCloseConnection.Enabled = True Then
            'Make sure that the database connection is
            'open before you try to add data back to
            'the controls.
            
            Call LoadDataInControls
            'reload the data.
        End If
    End If
End Sub

'***********
'*17th Step*
'***********

Private Sub WriteDataFromControls()
    
    
    'Again, there are several ways to manipulate field
    'values.  Some are shown.
    
    rsRecordSet("Login").Value = txtFirstName.Text
    
    rsRecordSet.Fields("Password1").Value = txtLastName.Text
    
    rsRecordSet!Level = txtTitle.Text
    
End Sub

'***********
'*18th Step*
'***********

Private Sub cmdSave_Click()
    
    If cmdAdd.Caption = "&Cancel" Then
        rsRecordSet.AddNew
    End If
    'This calls the Recordset MoveComplete event.
    'Before cmdEdit was added, there was no if-then statement.
    
    Call WriteDataFromControls
    'Write the data from the text boxes to the appropriate
    'fields.
    
    rsRecordSet.Update
    'No data is saved until the update method is executed.
    
    mblnAddMode = False
    'Saving closes add mode.
    
    cmdSave.Enabled = False
    'Save does not need to be enabled after the save is executed.
    
    If cmdAdd.Caption = "&Cancel" Then
        cmdAdd.Caption = "&Add"
    End If
    
    
    If cmdEdit.Caption = "&Cancel" Then
        cmdEdit.Caption = "&Edit"
    End If
    'Changes captions back to original.
    
    'These three steps prevent the user from adding
    'blank data or accidently changing data.
    txtFirstName.Locked = True
    txtLastName.Locked = True
    txtTitle.Locked = True
    
    Call EnableNavigation
    'Allow movement again.
    
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    'Add later to provide stability and consistency.
    
    rsRecordSet.Close
    rsRecordSet.Open
    'This is not a very elegant solution to a problem I have
    'with this database.  If I add a new record and then
    'immediately delete it, it reappears the next time the
    'database is opened.  If anyone has a slicker solution
    'let me know.
    
    lblEOF = rsRecordSet.EOF
    lblBOF = rsRecordSet.BOF
End Sub

'***********
'*19th Step* Add cmdDelete to the form.
'***********

'***********
'*20th Step*
'***********

Private Sub cmdDelete_Click()

    If rsRecordSet.EOF = False And _
        rsRecordSet.BOF = False And _
        cmdCloseConnection.Enabled = True Then
        'Check to see if there is data in the database
        'and make sure it is open.

        On Error Resume Next
        'If there is an error, ignore it.
        
        adoDataConn.begtrans
        'Deleting a record is important so the begtrans method
        'is used.  It makes sure all actions between begtrans
        'and committrans are done at the same time.
        rsRecordSet.Delete
        'Delete the record.
        
        adoDataConn.CommitTrans
        'The actions have been committed.
        
        rsRecordSet.MoveNext
        If rsRecordSet.EOF = True Then
            rsRecordSet.MoveLast
            'If the user deletes the record in the last position
            'go to the new record in the last position.
            
            If rsRecordSet.BOF = True Then
                Call ClearControls
                'If the last record is deleted, clear the text
                'boxes.
                
                MsgBox "There is no data in the recordset!", , "Oops!"
                'Alert the user that there is no more data in
                'the database.
            End If
        End If
    ElseIf rsRecordSet.EOF = True And rsRecordSet.BOF = True Then
        'Warn the user that he or she is trying to delete data
        'from a database with no records.
        
        MsgBox "There is no data in the recordset!", , "Oops!"
    End If
    
    lblBOF.Caption = rsRecordSet.BOF
    lblEOF.Caption = rsRecordSet.EOF
End Sub

'***********
'*21st Step*  Add cmdEdit to the form and code its click event.
'***********

Private Sub cmdEdit_Click()
    'This sub is very similar to cmdAdd.
    
    '********Some changes were made to cmdSave in order to
    '********allow edits to be saved.
    
    If cmdEdit.Caption = "&Edit" And _
        cmdCloseConnection.Enabled = True Then
        'Allow adds when the connection is opened.
            
        cmdEdit.Caption = "&Cancel"
        'Change the caption to Cancel and use this button
        'to prevent a new record from being saved.
    
        cmdSave.Enabled = True
        'Since a new record is being added, the user
        'should be allowed to save the data.
        
        Call DisableNavigation
        'No moves during edit.
        
        cmdAdd.Enabled = False
        cmdDelete.Enabled = False
        'Coded later to add more stability.
        
        txtFirstName.Locked = False
        txtLastName.Locked = False
        txtTitle.Locked = False
        'Allow the user to enter items into the text boxes.
        
        txtFirstName.SetFocus
        'Go to the first field.
        
    ElseIf cmdEdit.Caption = "&Cancel" Then
        cmdEdit.Caption = "&Edit"
        'If a user cancels, allow another add.
        
        cmdSave.Enabled = False
        'There will be no record to save.
        
        Call EnableNavigation
        'Allow movement.
        
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        'Coded later to add more stability.
        
        txtFirstName.Locked = True
        txtLastName.Locked = True
        txtTitle.Locked = True
        'Allow the user to enter items into the text boxes.
        
        If cmdCloseConnection.Enabled = True Then
            'Make sure that the database connection is
            'open before you try to add data back to
            'the controls.
            
            Call LoadDataInControls
            'reload the data.
        End If
    End If
End Sub

'***********
'*22nd Step*  Add cmdExit to the form and code its click event.
'***********

Private Sub cmdExit_Click()
    
    If cmdCloseConnection.Enabled = True Then
        'If the connection is opend, close it.
        Call cmdCloseConnection_Click
    End If
    
    Unload Me
    'Remove the form from memory.
    
    End
    'Terminate the project.
End Sub

'Summary.
'
'I.     Check Microsoft ActiveX DataObject 2.0 Library from the
'       Project Menu's Reference Menu.
'
'II.    A.  Declare your connection variarable.
'       B.  Declare your recordset variable.
'       C.  Declare a module level addstate boolean variable.
'
'III.   A.  Decide where to locate the code to open the
'           connection.
'           1.  Declare a connection string
'               a.  The string must have a provider.
'               b.  The string must have a path
'               c.  The string must have a database's name
'           2.  Set the connection object equal to a new ADODB
'               connection.
'           3.  Set the connection object's cursor location.
'           4.  Open the connection string with the connection
'               object.
'           5.  Open the connection object.
'           6.  Set the recordset object equal to a new ADODB
'               recordset.
'           7.  Set the recordset's properties.
'               a.  Select a cursor type.
'               b.  Select a curor location.
'               c.  Select a lock type.
'               d.  Select a source.
'               e.  Select a connection object.
'           8.  Open the recordset.
'       B.  Decide where to locate the code to close the
'           connection.
'           1.  Close the connection object.
'           2.  Set the connection object equal to nothing.
'
'IV.    Create controls on the form to view and enter data.
'
'V.     Create subs to load and clear controls holding data.
'
'VI.    Program the MoveComplete event to load data if the
'       program is not in add state.
'
'VII.   Provide recordset navigationi.
'       A.  Add command buttons for MoveFirst, MoveLast,
'           Previous, and Next.
'       B.  Check for BOF and EOF on Previous and Next.
'
'VIII.  Provide the ability to add records.
'       A.  Add a command button for adding records.
'       B.  Provide a method to cancel the current add.
'       C.  Add a command button that allows the user
'           to save a new addition.
'       D.  Code a procedure to move data from the form
'           to the database.
'
'IX.    Provide the user the ability to delete records.
'       A.  Add a command button for deleting records.
'       B.  Check for records before allowing deletes.
'
'X.     Provide the user the ability to edit records.
'       A.  Add a command button for editing recors.
'       B.  Allow the user to cancel the edit.
'       C.  Allow the user to save the changes.

'****************Extra Procedures for convience only************
Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
