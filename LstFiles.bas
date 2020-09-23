Attribute VB_Name = "LstFiles"
Public MRUX(15)

Public Function FormatFileSize(Size As String) As String


    Dim strSize As String, strSize2 As String, strSize3 As String, strSize4 As String
    Dim FormattedSize As String
    Dim i As Integer
    ' if the size is greater than 999, insert a comma
    ' 3 spaces from the end or 3 space from the last
    ' comma
    ' convert the "number" to a string
    strSize = Str(Size)
    
    ' rebuild the string adding commas


    Select Case Len(strSize)
        Case 5
        ' format 1,000
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 2)
        strSize3 = strSize3 + strSize2 + " Kb"
        Case 6
        ' format 10,000
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 3)
        strSize3 = strSize3 + strSize2 + " Kb"
        Case 7
        ' format 100,000
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 4)
        strSize3 = strSize3 + strSize2 + " Kb"
        Case 8
        ' format 1,000,000
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 5)
        strSize3 = strSize3 + strSize2
        
        strSize4 = Right(strSize3, 7)
        strSize4 = "," + strSize4
        strSize3 = Left(strSize, 2)
        strSize3 = strSize3 + strSize4 + " Mb"
        Case 9
        ' format 10,000,000
        strSize2 = Right(strSize, 3)
        strSize2 = "," + strSize2
        strSize3 = Left(strSize, 5)
        strSize3 = strSize3 + strSize2
        
        strSize4 = Right(strSize3, 7)
        strSize4 = "," + strSize4
        strSize3 = Left(strSize, 3)
        strSize3 = strSize3 + strSize4 + " Mb"
        Case Else
        strSize3 = strSize + " bytes"
    End Select




FormattedSize = strSize3


    FormatFileSize = FormattedSize

End Function

