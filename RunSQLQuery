Sub Run_Query()

    Dim Server_Name, Database_Name, User_ID, Password, Driver_Name, SQLA, SQLB As String
    Dim OutputArray()
    Dim OutputValue as String
        
    Set rs = CreateObject("ADODB.Recordset")
    Server_Name = ""
    Database_Name = ""
    User_ID = ""
    Password = ""
    Driver_Name = ""

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Driver = " & Driver_Name & _
      ";Server=" & Server_Name & _
      ";Database=" & Database_Name & _
      ";Uid=" & User_ID & _
      ";Pwd=" & Password & ";"

    SQLA = "Select * from Table"
    rs.Open SQLA, Cn

    SQLB = "Select max(col) from Table"
    rs.Open SQLB, Cn
    OutputArray = rs.GetRows()
    OutputValue = OutputArray(0,0)

    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing

End Sub
