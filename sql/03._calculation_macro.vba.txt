Sub CalculateAndScoreLead()
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim clientName As String
    Dim routeName As String
    Dim containerType As String
    Dim margin As Double
    Dim baseCost As Double
    Dim finalPrice As Double
    
    clientName = Range("C3").Value
    routeName = Range("C4").Value
    containerType = Range("C5").Value
    margin = Val(Range("C6").Value)
    
    If routeName = "" Or containerType = "" Then
        MsgBox "Пожалуйста, выберите маршрут и тип контейнера!", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    strConn = "DSN=my_freight_db;"
              
    Set conn = CreateObject("ADODB.Connection")
    On Error GoTo ErrorHandler
    conn.Open strConn
    
    strSQL = "SELECT base_cost_usd FROM freight_tariffs WHERE route_name = '" & routeName & "' AND container_type = '" & containerType & "';"
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conn
    
    If Not rs.EOF Then
        baseCost = rs.Fields("base_cost_usd").Value
    Else
        MsgBox "Тариф не найден в базе данных MySQL!", vbCritical, "Ошибка"
        rs.Close: conn.Close
        Exit Sub
    End If
    
    rs.Close
    conn.Close
    
    finalPrice = baseCost + margin
    
    Range("F4").Value = baseCost
    Range("F5").Value = finalPrice
    Range("F4").NumberFormat = "[$$-409]#,##0"
    Range("F5").NumberFormat = "[$$-409]#,##0"

    If margin >= 500 Then
        Range("F6").Value = "ВЫСОКАЯ ПРИБЫЛЬ"
        Range("F6").Font.Color = RGB(0, 128, 0)
        Range("F6").Font.Bold = True
    Else
        Range("F6").Value = "НИЗКАЯ ПРИБЫЛЬ"
        Range("F6").Font.Color = RGB(128, 0, 0)
        Range("F6").Font.Bold = True
    End If

    Dim pdfPath As String
    pdfPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Freight_Quotation_" & clientName & ".pdf"
    
    On Error Resume Next
    Kill pdfPath
    On Error GoTo 0
    
    On Error Resume Next
    Range("E5:G5").ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=True, OpenAfterPublish:=False
    On Error GoTo ErrorHandler
    
    MsgBox "Расчет окончен! Официальное КП для клиента сохранено на рабочий стол.", vbInformation, "Успех"

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка подключения к MySQL: " & Err.Description, vbCritical, "Технический сбой"
End Sub

