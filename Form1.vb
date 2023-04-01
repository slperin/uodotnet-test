Imports System.Threading
Imports IBMU2.UODOTNET

Public Class FrmMain

  'TestU2
  'Test UniFile class
  'Baseware Systems
  'Ing. Sergio Perin (slperin@baseware.com.ar)

  Dim DBConn As UniSession
  Dim DBFile As UniFile
  Dim DBRecord As UniDynArray
  Dim UVFileEx As UniFileException

  Dim ItemID As String
  Dim ReaduFlag As Boolean
  Dim ReleaseFlag As Boolean

  Dim SrvName As String = ""
  Dim AccName As String = ""
  Dim UsrName As String = ""
  Dim Password As String = ""

  Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

    Label1.Text = ""
    Label2.Text = ""

    Try
      DBConn = UniObjects.OpenSession(SrvName, UsrName, Password, AccName, "uvcs")
      If DBConn Is Nothing Then
        Label1.Text = "Connection failed with server " & SrvName & "."
      Else
        Label1.Text = "Connection established with server " & SrvName & "." & vbCrLf
        Label1.Text &= "Server " & IIf(DBConn.Service = "uvcs", "UniVerse", "UniData").ToString
        DBFile = DBConn.CreateUniFile("VOC")
        DBFile.UniFileBlockingStrategy = UniObjectsTokens.RETURN_ON_LOCKED
        DBFile.UniFileLockStrategy = UniObjectsTokens.NO_LOCKS
        DBFile.UniFileReleaseStrategy = UniObjectsTokens.UVT_EXPLICIT_RELEASE
        DBFile.RecordID = "RELLEVEL"
        DBFile.ReadField(2)
        DBRecord = DBFile.Record
        Label1.Text &= " Version " & DBRecord.StringValue
      End If
    Catch UVFileEx As UniFileException
      Label1.Text = UVFileEx.Message
    Catch Ex As Exception
      Label1.Text = Ex.Message
    End Try

  End Sub

  Private Sub FrmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

    If DBConn IsNot Nothing Then
      If DBConn.IsActive Then
        UniObjects.CloseSession(DBConn)
      End If
    End If

  End Sub

  Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    ItemID = "LIXTU" 'Any item that don't exists in the VOC file.
    ReaduFlag = False
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC without locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be missing on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30001." & vbCrLf
    Label1.Text &= "Previous to this, ItemID " & ItemID & " on file VOC must not be locked by any other process."
    Label2.Text = ""

    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."

  End Sub

  Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    ItemID = "LIXTU" 'Any item that don't exists in the VOC file.
    ReaduFlag = True
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be missing on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30001." & vbCrLf
    Label1.Text &= "Previous to this, ItemID " & ItemID & " on file VOC must not be locked by any other process." & vbCrLf
    Label1.Text &= "ItemID not released after each read."
    Label2.Text = ""

    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST FAILED."

  End Sub

  Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    ItemID = "LIXTU" 'Any item that don't exists in the VOC file.
    ReaduFlag = True
    ReleaseFlag = True

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be missing on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30001." & vbCrLf
    Label1.Text &= "Previous to this, ItemID " & ItemID & " on file VOC must not be locked by any other process." & vbCrLf
    Label1.Text &= "ItemID released after each read."
    Label2.Text = ""

    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."

  End Sub

  Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    ItemID = "LIXTU" 'Any item that don't exists in the VOC file.
    ReaduFlag = False
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC without locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be missing on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30001." & vbCrLf
    Label1.Text &= "Previous to this, you must lock record " & ItemID & " on file VOC."
    Label2.Text = ""

    MsgBox("Lock the item " & ItemID & " on another process and press OK button to run test.")
    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."
    MsgBox("Release the item " & ItemID & " previously locked and press OK button to continue.")

  End Sub

  Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    ItemID = "LIXTU" 'Any item that don't exists in the VOC file.
    ReaduFlag = True
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be missing on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30002." & vbCrLf
    Label1.Text &= "Previous to this, you must lock record " & ItemID & " on file VOC."
    Label2.Text = ""

    MsgBox("Lock the item " & ItemID & " on another process and press OK button to run test.")
    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."
    MsgBox("Release the item " & ItemID & " previously locked and press OK button to continue.")

  End Sub

  Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    ItemID = "LISTU" 'Any item that must exists in the VOC file.
    ReaduFlag = False
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC without locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be present on file VOC." & vbCrLf
    Label1.Text &= "Previous to this, ItemID " & ItemID & " on file VOC must not be locked by any other process." & vbCrLf
    Label1.Text &= "Must show the record body without error code."
    Label2.Text = ""

    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."

  End Sub

  Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    ItemID = "LISTU" 'Any item that must exists in the VOC file.
    ReaduFlag = True
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be present on file VOC." & vbCrLf
    Label1.Text &= "Must show the record body without error code." & vbCrLf
    Label1.Text &= "Previous to this, ItemID " & ItemID & " on file VOC must not be locked by any other process." & vbCrLf
    Label1.Text &= "ItemID not released after each read. "
    Label2.Text = ""

    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."

  End Sub

  Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    ItemID = "LISTU" 'Any item that must exists in the VOC file.
    ReaduFlag = True
    ReleaseFlag = True

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be present on file VOC." & vbCrLf
    Label1.Text &= "Must show the record body without error code." & vbCrLf
    Label1.Text &= "Previous to this, ItemID " & ItemID & " on file VOC must not be locked by any other process." & vbCrLf
    Label1.Text &= "ItemID released after each read. "
    Label2.Text = ""

    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."

  End Sub

  Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

    ItemID = "LISTU" 'Any item that must exists in the VOC file.
    ReaduFlag = False
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be present on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30001." & vbCrLf
    Label1.Text &= "Previous to this, you must lock record " & ItemID & " on file VOC."
    Label2.Text = ""

    MsgBox("Lock the item " & ItemID & " on another process and press OK button to run test.")
    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."
    MsgBox("Release the item " & ItemID & " previously locked and press OK button to continue.")

  End Sub

  Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

    ItemID = "LISTU" 'Any item that must exists in the VOC file.
    ReaduFlag = True
    ReleaseFlag = False

    Label1.Text = "Three attempts to read record " & ItemID & " on file VOC with locking." & vbCrLf
    Label1.Text &= "ItemID " & ItemID & " must be present on file VOC." & vbCrLf
    Label1.Text &= "Must show the exception with error code 30002." & vbCrLf
    Label1.Text &= "Previous to this, you must lock record " & ItemID & " on file VOC." & vbCrLf
    Label1.Text &= "Release item after the test is completed."
    Label2.Text = ""

    MsgBox("Lock the item " & ItemID & " on another process and press OK button to run test.")
    Test(ItemID)
    Label2.Text &= vbCrLf & "TEST PASSED."
    MsgBox("Release the item " & ItemID & " previously locked and press OK button to continue.")

  End Sub

  Private Sub Test(ByVal ItemID As String)

    If ReaduFlag Then
      DBFile.UniFileLockStrategy = UniObjectsTokens.UVT_EXCLUSIVE_READ
    Else
      DBFile.UniFileLockStrategy = UniObjectsTokens.NO_LOCKS
    End If
    DBFile.RecordID = ItemID

    Try
      DBFile.Read()
      DBRecord = DBFile.Record
      Label2.Text &= "1 -> No thrown exception. Record " & ItemID & " is [" & DBRecord.StringValue & "]." & vbCrLf
    Catch UVFileEx As UniFileException
      Select Case UVFileEx.ErrorCode
        Case 30001
          Label2.Text &= "1 -> Record not on file (ERRORCODE=30001)" & vbCrLf
        Case 30002
          Label2.Text &= "1 -> Record locked (ERRORCODE=30002)" & vbCrLf
        Case Else
          Label2.Text &= "1 -> Exception: " & UVFileEx.Message & vbCrLf
      End Select
    Catch Ex As Exception
      Label2.Text &= "1 -> Exception: " & Ex.Message & vbCrLf
    End Try
    If ReleaseFlag Then DBFile.UnlockRecord()

    Try
      DBFile.Read()
      DBRecord = DBFile.Record
      Label2.Text &= "2 -> No thrown exception. Record " & ItemID & " is [" & DBRecord.StringValue & "]." & vbCrLf
    Catch UVFileEx As UniFileException
      Select Case UVFileEx.ErrorCode
        Case 30001
          Label2.Text &= "2 -> Record not on file (ERRORCODE=30001)" & vbCrLf
        Case 30002
          Label2.Text &= "2 -> Record locked (ERRORCODE=30002)" & vbCrLf
        Case Else
          Label2.Text &= "2 -> Exception: " & UVFileEx.Message & vbCrLf
      End Select
    Catch Ex As Exception
      Label2.Text &= "2 -> Exception: " & Ex.Message & vbCrLf
    End Try
    If ReleaseFlag Then DBFile.UnlockRecord()

    Try
      DBFile.Read()
      DBRecord = DBFile.Record
      Label2.Text &= "3 -> No thrown exception. Record " & ItemID & " is [" & DBRecord.StringValue & "]." & vbCrLf
    Catch UVFileEx As UniFileException
      Select Case UVFileEx.ErrorCode
        Case 30001
          Label2.Text &= "3 -> Record not on file (ERRORCODE=30001)" & vbCrLf
        Case 30002
          Label2.Text &= "3 -> Record locked (ERRORCODE=30002)" & vbCrLf
        Case Else
          Label2.Text &= "3 -> Exception: " & UVFileEx.Message & vbCrLf
      End Select
    Catch Ex As Exception
      Label2.Text &= "3 -> Exception: " & Ex.Message & vbCrLf
    End Try
    DBFile.UnlockRecord()

  End Sub

End Class