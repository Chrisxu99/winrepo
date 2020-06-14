VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetInquiryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub IQ_FDC()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim Sender, sentDate As String
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@fdclabs.com") > 0 And InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If
                    Set objRows = html.getElementsbyTagName("tr")

                    If InStr(MailText, "sorry it is unavailable") + InStr(MailText, "没法提供") + InStr(MailText, "放弃报价") _
                     + InStr(MailText, "已停用") > 0 Then
                       
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "FDC"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Partner Cat#") < 150 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If InStr(strObj, "Terms") > 0 Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 5 Then k = 5
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  'If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                  If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If
                                  

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""))
                                  Font = ""
                                  Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  On Error Resume Next
                                   Font = objFont(0).Color
                                  On Error GoTo 0
 
                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "FDC"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     objMail.UnRead = False
                                     m = m + 1
                                     If Font = "#ff0000" Then
                                        .Cells(r1 + 1, 28).Value = "Change the Purity"
                                     End If
                                     For i1 = 1 To k
                                         Price = ""
                                         UPrice = ""
                                         ShipPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = UCase(Trim(Replace(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", ""), "$", "")))
                                         strLT = UCase(Trim(Replace(objRows(j1 + 4 + i1).Cells(2).innerText, " ", "")))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               m2 = InStr(strPrice, "+")
                                               m3 = InStr(strPrice, "SHIP")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               If m1 > 0 And m2 = 0 And m3 = 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                               Else
                                                  If m1 > 0 And m2 > 0 And m3 > 0 Then
                                                     UPrice = Trim(Left(strPrice, m1 - 1))
                                                     PriceUnit = Trim(Mid(strPrice, m1 + 1, m2 - m1 - 1))
                                                     ShipPrice = Trim(Mid(strPrice, m2 + 1, m3 - m2 - 1))
                                                  Else
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     If ShipPrice = "" Then
                                                        Price = CStr(QtyNum * UPrice)
                                                     Else
                                                        Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                     End If
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     If ShipPrice = "" Then
                                                        Price = CStr(QtyNum * UPrice)
                                                     Else
                                                        Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                     End If
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     If ShipPrice = "" Then
                                                        Price = CStr(QtyNum * UPrice * 1000)
                                                     Else
                                                        Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                     End If
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            m8 = InStr(strLT, "WEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "DAY")
                                            'If m8 = 0 Then m8 = InStr(strLT, "DYA")
                                            m9 = InStr(strLT, ",")
                                            If m9 = 0 Then m9 = InStr(strLT, "%")
                                            If m9 > 0 And m8 > 0 Then
                                               If m9 > m8 Then
                                                  strLT = Trim(Left(strLT, m9 - 1))
                                               Else
                                                  strLT = Trim(Right(strLT, Len(strLT) - m9))
                                               End If
                                            End If
                                            
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WEEK")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                         'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                         If Price <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = GetNum(LT)
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\FDC_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("FDC Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from FDC have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI
 
CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub

Public Sub IQ_INFOARK()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@infoark.com.cn") > 0 And _
                (InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Or InStr(sTitle, "Re: 转发: Inquiry from AstaTech") = 1) Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If

                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "提供不了") + InStr(MailText, "不提供") + InStr(MailText, "不报价") + InStr(MailText, "无法提供") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "no stock") _
                    + InStr(Replace(MailText, " ", ""), "can'tsupply") + InStr(MailText, "报不了") + InStr(MailText, "暂无报价") > 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "INFOARK"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Partner Cat#") < 100 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0

                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  'strCmnts = objRows(j2 + 3).Cells(0).innerText
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "INFOARK"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = UCase(Trim(xlApp.WorksheetFunction.Clean(Replace(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", ""), "$", ""))))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               m2 = InStr(strPrice, "BY")
                                               m3 = InStr(strPrice, "(")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               If m1 > 0 And m2 = 0 And m3 = 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                               Else
                                                  If m1 > 0 And m2 > 0 And m3 = 0 Then
                                                     UPrice = Trim(Left(strPrice, m1 - 1))
                                                     PriceUnit = Trim(Mid(strPrice, m1 + 1, m2 - m1 - 1))
                                                     'ShipPrice = Trim(Mid(strPrice, m2 + 1, m3 - m2 - 1))
                                                  Else
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "WEEL")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            .Cells(r1 + 1, 7 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            If Price <> "" Then .Cells(r1 + 1, 8 + i1 * 3).Value = LT
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            If Price <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = LT
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\INFOARK_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("INFOARK Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from INFOARK have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub

Public Sub IQ_PHARMABLOCK()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@PharmaBlock.com") > 0 And _
                (InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Or InStr(sTitle, "答复: Inquiry from AstaTech") = 1) Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText

                    n1 = InStr(MailText, "发件人:")
                    n2 = InStr(MailText, "发送时间:")
                    n3 = InStr(MailText, "收件人:")
                    If n1 = 0 Then n1 = InStr(MailText, "From:")
                    If n2 = 0 Then n2 = InStr(MailText, "Sent:")
                    If n3 = 0 Then n3 = InStr(MailText, "To:")
                    If n1 > 0 And n2 > n1 And n3 > n2 And (n2 - n1) < 100 And (n3 - n2) < 100 Then
                       strSender = Trim(Mid(MailText, n1, n2 - n1 - 4))
                       strDate = Trim(Mid(MailText, n2 + 5, n3 - n2 - 5))
                       n4 = InStr(strDate, "年")
                       n5 = InStr(strDate, "月")
                       n6 = InStr(strDate, "日")
                       If n4 > 0 And n5 > 0 And n6 > 0 Then
                       sentDate = DateSerial(Left(strDate, n4 - 1), Mid(strDate, n4 + 1, n5 - n4 - 1), Mid(strDate, n5 + 1, n6 - n5 - 1))
                       End If
                       Sender = Trim(Right(strSender, Len(strSender) - InStrRev(strSender, ":")))
                    Else
                       GoTo NextI
                    End If
                    
                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "提供不了") + InStr(MailText, "不提供") + InStr(MailText, "不报价") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "will not provide") > 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "PHARMABLOCK"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Partner Cat#") < 230 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then
                                 j1 = j
                            'If strObj = "Extra Quoted Quantity" Then
                               'j2 = j
                               'k = j2 - j1 - 5
                               'If k > 3 Then k = 3
                               'If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0
                                  
                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  'strCmnts = objRows(j2 + 3).innerText
                                  'Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  'If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Then
                                     'GoTo NextI
                                  'End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "PHARMABLOCK"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     '.Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         'If i1 = k + 1 Then
                                            'GoTo Skip
                                         'End If
                                         On Error Resume Next
                                         strQty = Trim(objRows(j1 + 4 + i1).Cells(0).innerText)
                                         On Error GoTo 0
                                         If strQty = "Partner Cat#:" Then
                                            GoTo NextI
                                         Else
                                            strQty = Trim(Replace(strQty, " ", ""))
                                            strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                            Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         End If
                                         On Error Resume Next
                                         strPrice = Trim(Replace(Replace(UCase(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", "")), "$", ""), "USD", ""))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         On Error GoTo 0
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               'm2 = InStr(strPrice, "BY")
                                               m3 = InStr(strPrice, "(")
                                               If m3 = 0 Then m3 = InStr(strPrice, "FINAL")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               If m3 > 0 Then strPrice = Trim(Left(strPrice, m3 - 1))
                                               If m1 > 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                                  If IsNumeric(UPrice) = False Or (PriceUnit <> "G" And PriceUnit <> "MG" And PriceUnit <> "KG") Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unforamtted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               Else
                                                  Price = strPrice
                                                  If Not IsNumeric(Price) Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WEEK")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         
                                         'If i1 > k + 1 Then
                                            '.Cells(r1 + 1, 4 + i1 * 3).Value = Qty
                                            '.Cells(r1 + 1, 5 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            'If Price <> "" Then .Cells(r1 + 1, 6 + i1 * 3).Value = LT
                                         'Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            If Price <> "" And IsNumeric(Price) Then
                                               .Cells(r1 + 1, 10 + i1 * 3).Value = CStr(Round(Price * 1, 0))
                                               .Cells(r1 + 1, 11 + i1 * 3).Value = LT
                                            End If
                                         'End If
'Skip:
                                     Next i1
                                  End With
                               'End If
                              Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\PHARMABLOCK_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("PHARMABLOCK Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from PHARMABLOCK have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub

Public Sub IQ_SUNWAY()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@sunwaypharm.com") > 0 And InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If
                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "提供不了") + InStr(MailText, "不提供") + InStr(MailText, "不报价") + InStr(MailText, "不能提供") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "donot have it") _
                    + InStr(MailText, "do not have it") + InStr(MailText, "don't have it") + InStr(MailText, "cannot make it") _
                    + InStr(MailText, "cannot supply") > 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "SUNWAY"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Partner Cat#") < 100 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0
                                  
                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Or _
                                     InStr(UCase(Rqmts), "HCL") + InStr(UCase(Rqmts), "HBR") + InStr(UCase(Rqmts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "SUNWAY"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                        On Error Resume Next
                                         strQty = Trim(objRows(j1 + 4 + i1).Cells(0).innerText)
                                         strQty = Trim(Replace(strQty, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = Trim(Replace(Replace(Replace(Replace(UCase(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", "")), "$", ""), "USD", ""), "UDS", ""), "SUD", ""))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                        On Error GoTo 0
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               'm2 = InStr(strPrice, "BY")
                                               'm3 = InStr(strPrice, "(")
                                               'If m3 = 0 Then m3 = InStr(strPrice, "FINAL")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR") _
                                                    + InStr(strLT, "HCL") + InStr(strLT, "SALT") + InStr(strLT, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               'If m3 > 0 Then strPrice = Trim(Left(strPrice, m3 - 1))
                                               If m1 > 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                                  If IsNumeric(UPrice) = False Or (PriceUnit <> "G" And PriceUnit <> "MG" And PriceUnit <> "KG") Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unforamtted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               Else
                                                  Price = strPrice
                                                  If Not IsNumeric(Price) Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                           On Error Resume Next
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            'm5 = InStr(strLT, "STOCK")
                                            'm6 = InStr(strLT, "-")
                                            'm7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "DAY")
                                            If m8 = 0 Then m8 = InStr(strLT, "DYA")
                                            If m8 = 0 Then m8 = InStr(strLT, "DASY")
                                            m9 = InStr(strLT, ",")
                                            If m9 > 0 And m8 > 0 Then
                                               If m9 > m8 Then
                                                  strLT = Trim(Left(strLT, m9 - 1))
                                               Else
                                                  strLT = Trim(Right(strLT, Len(strLT) - m9))
                                               End If
                                            End If
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            If m7 = 0 Then m7 = InStr(strLT, "DYA")
                                            If m7 = 0 Then m7 = InStr(strLT, "DASY")
                                            m8 = InStr(strLT, "WEEK")
                                           On Error GoTo 0
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 7 + i1 * 3).Value = Price
                                               .Cells(r1 + 1, 8 + i1 * 3).Value = LT
                                            End If
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                               .Cells(r1 + 1, 11 + i1 * 3).Value = LT
                                            End If
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\SUNWAY_InquiryInfo.xlsx"
       
      On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
      On Error GoTo 0
      
       'WB1.Close
       MsgBox ("SUNWAY Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from SUNWAY have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub

Public Sub IQ_FEIBO()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
 
 If MsgBox("Please make sure to set up the automatic forward rules for" & vbNewLine _
    & "FEIBO Inquiry replies in Outlook before using this program." & vbNewLine _
    & "Proceed or Not?", vbYesNo) = vbNo Then
    GoTo CleanUp
 End If
 
 '============================================================================================
 'File - Info - Manage Rules&Alerts - New Rule - with specific words in the sender's address &
 'with specific words in the subject - mark it as read & forward it to people or public group
 '============================================================================================
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 fAddress = objMail.SenderEmailAddress
                 fTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(fAddress, "syi@astatechinc.com") > 0 And InStr(fTitle, "FW: Inquiry from AstaTech") = 1 Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                   On Error Resume Next
                    n0 = InStr(MailText, "Web:")
                    n1 = InStr(n0, MailText, "在")
                    n2 = InStr(n1, MailText, "年")
                    n3 = InStr(n1, MailText, "月")
                    n4 = InStr(n3, MailText, "<")
                    n5 = InStr(n3, MailText, ">")
                   On Error GoTo 0

                    If n0 > 0 And n1 > 0 And n2 > 0 And n3 > 0 And n4 > 0 And n5 > 0 Then
                       Sender = Trim(Mid(MailText, n4 + 1, n5 - n4 - 1))
                       strDate = DateSerial(Trim(Mid(MailText, n1 + 1, n2 - n1 - 1)), Trim(Mid(MailText, n2 + 1, n3 - n2 - 1)), Trim(Mid(MailText, n3 + 1, 2)))
                       sentDate = Format(strDate, "mm/dd/yyyy")
                    Else
                       GoTo NextI
                    End If
                    
                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "cannot be synthesized") + InStr(MailText, "can't be synthesized") + InStr(MailText, "can't make a new batch") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "check with other vendors") _
                    + InStr(MailText, "out of stock") + InStr(MailText, "can't supply") > 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "FEIBO"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Partner Cat#") < 500 Then
                          compID = Trim(Right(fTitle, Len(fTitle) - InStr(fTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0
                                  
                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "FEIBO"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                         On Error Resume Next
                                         strQty = Trim(objRows(j1 + 4 + i1).Cells(0).innerText)
                                         On Error GoTo 0
                                         strQty = Trim(Replace(strQty, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = Trim(Replace(Replace(UCase(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", "")), "$", ""), "USD", ""))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               'm2 = InStr(strPrice, "BY")
                                               'm3 = InStr(strPrice, "(")
                                               'If m3 = 0 Then m3 = InStr(strPrice, "FINAL")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               'If m3 > 0 Then strPrice = Trim(Left(strPrice, m3 - 1))
                                               If m1 > 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                                  If IsNumeric(UPrice) = False Or (PriceUnit <> "G" And PriceUnit <> "MG" And PriceUnit <> "KG") Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unforamtted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               Else
                                                  Price = strPrice
                                                  If Not IsNumeric(Price) Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            'm5 = InStr(strLT, "STOCK")
                                            'm6 = InStr(strLT, "-")
                                            'm7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "DAY")
                                            If m8 = 0 Then m8 = InStr(strLT, "DYA")
                                            m9 = InStr(strLT, ",")
                                            If m9 > 0 And m8 > 0 Then
                                               If m9 > m8 Then
                                                  strLT = Trim(Left(strLT, m9 - 1))
                                               Else
                                                  strLT = Trim(Right(strLT, Len(strLT) - m9))
                                               End If
                                            End If
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "~")
                                            If m6 = 0 Then m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            If m7 = 0 Then m7 = InStr(strLT, "DYA")
                                            m8 = InStr(strLT, "WEEK")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 7 + i1 * 3).Value = Price
                                               .Cells(r1 + 1, 8 + i1 * 3).Value = GetNum(LT)
                                            End If
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                               .Cells(r1 + 1, 11 + i1 * 3).Value = GetNum(LT)
                                            End If
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\FEIBO_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("FEIBO Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from FEIBO have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub


Function GetNum(ByVal strNum As String) As String

   strNum = UCase(Trim(strNum))
   Select Case strNum
       Case "ONE"
          GetNum = "1"
       Case "TWO"
          GetNum = "2"
       Case "THREE"
          GetNum = "3"
       Case "FOUR"
          GetNum = "4"
       Case "FIVE"
          GetNum = "5"
       Case "SIX"
          GetNum = "6"
       Case "SEVEN"
          GetNum = "7"
       Case "EIGHT"
          GetNum = "8"
       Case "NINE"
          GetNum = "9"
       Case "TEN"
          GetNum = "10"
       Case "ELEVEN"
          GetNum = "11"
       Case "TWELVE"
          GetNum = "12"
       Case "THIRTEEN"
          GetNum = "13"
       Case "FOURTEEN"
          GetNum = "14"
       Case Else
          GetNum = strNum
    End Select

End Function

Public Sub IQ_YINGXI()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@outlook.com") > 0 And InStr(sTitle, "答复: Inquiry from AstaTech") = 1 Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                   On Error Resume Next
                    n0 = InStr(MailText, "Terms:")
                    n1 = InStr(n0, MailText, "发件人:")
                    n2 = InStr(n1, MailText, "<")
                    n3 = InStr(n2, MailText, ">")
                    n4 = InStr(n3, MailText, "发送时间:")
                    n5 = InStr(n4, MailText, "年")
                    n6 = InStr(n5, MailText, "月")
                    n7 = InStr(n6, MailText, "日")
                   On Error GoTo 0
                    
                    If n0 > 0 And n1 > 0 And n4 > 0 And n5 > 0 Then
                       Sender = Trim(Mid(MailText, n2 + 1, n3 - n2 - 1))
                       strDate = Trim(Mid(MailText, n4 + 5, n7 - n4 - 4))
                       s1 = InStr(strDate, "年")
                       s2 = InStr(strDate, "月")
                       s3 = InStr(strDate, "日")
                       sentDate = DateSerial(Trim(Left(strDate, s1 - 1)), Trim(Mid(strDate, s1 + 1, s2 - s1 - 1)), Trim(Mid(strDate, s2 + 1, s3 - s2 - 1)))
                    Else
                       GoTo NextI
                    End If
                    
                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "提供不了") + InStr(MailText, "不提供") + InStr(MailText, "不报价") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "can't provide") > 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "YINGXI"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Partner Cat#") < 150 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""), "purity", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0
                                  
                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Or _
                                     InStr(UCase(Rqmts), "HCL") + InStr(UCase(Rqmts), "HBR") + InStr(UCase(Rqmts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "YINGXI"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                         On Error Resume Next
                                         strQty = Trim(objRows(j1 + 4 + i1).Cells(0).innerText)
                                         On Error GoTo 0
                                         strQty = Trim(Replace(strQty, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = Trim(Replace(Replace(UCase(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", "")), "$", ""), "USD", ""))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               'm2 = InStr(strPrice, "BY")
                                               'm3 = InStr(strPrice, "(")
                                               'If m3 = 0 Then m3 = InStr(strPrice, "FINAL")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               'If m3 > 0 Then strPrice = Trim(Left(strPrice, m3 - 1))
                                               If m1 > 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                                  If IsNumeric(UPrice) = False Or (PriceUnit <> "G" And PriceUnit <> "MG" And PriceUnit <> "KG") Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unforamtted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               Else
                                                  Price = strPrice
                                                  If Not IsNumeric(Price) Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            'm5 = InStr(strLT, "STOCK")
                                            'm6 = InStr(strLT, "-")
                                            'm7 = InStr(strLT, "DAY")
                                            'm8 = InStr(strLT, "WEEK")
                                            'If m8 = 0 Then m8 = InStr(strLT, "DAY")
                                            'If m8 = 0 Then m8 = InStr(strLT, "DYA")
                                            'm9 = InStr(strLT, ",")
                                            'If m9 > 0 And m8 > 0 Then
                                               'If m9 > m8 Then
                                                  'strLT = Trim(Left(strLT, m9 - 1))
                                               'Else
                                                  'strLT = Trim(Right(strLT, Len(strLT) - m9))
                                               'End If
                                            'End If
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            If m7 = 0 Then m7 = InStr(strLT, "DYA")
                                            m8 = InStr(strLT, "WEEK")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 7 + i1 * 3).Value = Price
                                               .Cells(r1 + 1, 8 + i1 * 3).Value = LT
                                            End If
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                               .Cells(r1 + 1, 11 + i1 * 3).Value = LT
                                            End If
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\YINGXI_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("YINGXI Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from YINGXI have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub


Public Sub IQ_WUXI()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@wuxiapptec.com") > 0 And (InStr(sTitle, "RE: Inquiry from AstaTech") = 1 Or _
                   InStr(sTitle, "答复: Inquiry from AstaTech") = 1) Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    
                   On Error Resume Next
                    n1 = InStr(MailText, "From:")
                    n2 = InStr(n1, MailText, "[")
                    n3 = InStr(n2, MailText, "]")
                    n4 = InStr(n3, MailText, "Sent:")
                    n5 = InStr(n4, MailText, "To:")
                   On Error GoTo 0
                    
                    If n1 > 0 And n2 > 0 And n3 > 0 And n4 > 0 And n5 > 0 Then
                       strSender = Trim(Mid(MailText, n2 + 1, n3 - n2 - 1))
                       n6 = InStr(strSender, ":")
                       Sender = Trim(Right(strSender, Len(strSender) - n6))
                       strDate = Trim(Mid(MailText, n4 + 5, n5 - n4 - 5))
                       n7 = InStr(strDate, ",")
                       n8 = InStrRev(strDate, ":")
                       sDate = Trim(Mid(strDate, n7 + 1, n8 - n7 - 3))
                       sentDate = Format(sDate, "mm/dd/yyyy")
                    Else
                      On Error Resume Next
                       p1 = InStr(MailText, "发件人:")
                       p2 = InStr(p1, MailText, "[")
                       p3 = InStr(p2, MailText, "]")
                       p4 = InStr(p3, MailText, "发送时间:")
                       p5 = InStr(p4, MailText, "年")
                       p6 = InStr(p5, MailText, "月")
                       p7 = InStr(p6, MailText, "日")
                      On Error GoTo 0
                    
                       If p1 > 0 And p2 > 0 And p4 > 0 And p5 > 0 Then
                          strSender = Trim(Mid(MailText, p2 + 1, p3 - p2 - 1))
                          If InStr(strSender, ":") > 0 Then Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                          strDate = Trim(Mid(MailText, p4 + 5, p7 - p4 - 4))
                          s1 = InStr(strDate, "年")
                          s2 = InStr(strDate, "月")
                          s3 = InStr(strDate, "日")
                          sentDate = DateSerial(Trim(Left(strDate, s1 - 1)), Trim(Mid(strDate, s1 + 1, s2 - s1 - 1)), Trim(Mid(strDate, s2 + 1, s3 - s2 - 1)))
                       Else
                          GoTo NextI
                       End If
                    End If
                    
                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "提供不了") + InStr(MailText, "不提供") + InStr(MailText, "不报价") + InStr(MailText, "can't offer") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "cannot provide") > 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "WUXI"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "find the quotation in you forms") + InStr(MailText, "Here comes our quotation as") _
                        + InStr(MailText, "Here is WuXi AppTec Catalog quotation") + InStr(MailText, "Below is our quotation") > 0 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""), "purity", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0
                                  
                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Or _
                                     InStr(UCase(Rqmts), "HCL") + InStr(UCase(Rqmts), "HBR") + InStr(UCase(Rqmts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "WUXI"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                         On Error Resume Next
                                         strQty = Trim(objRows(j1 + 4 + i1).Cells(0).innerText)
                                         On Error GoTo 0
                                         strQty = Trim(Replace(strQty, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = Trim(Replace(Replace(UCase(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", "")), "$", ""), "USD", ""))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               'm2 = InStr(strPrice, "BY")
                                               'm3 = InStr(strPrice, "(")
                                               'If m3 = 0 Then m3 = InStr(strPrice, "FINAL")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               'If m3 > 0 Then strPrice = Trim(Left(strPrice, m3 - 1))
                                               If m1 > 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                                  If IsNumeric(UPrice) = False Or (PriceUnit <> "G" And PriceUnit <> "MG" And PriceUnit <> "KG") Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unforamtted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               Else
                                                  Price = strPrice
                                                  If Not IsNumeric(Price) Then
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            'm5 = InStr(strLT, "STOCK")
                                            'm6 = InStr(strLT, "-")
                                            'm7 = InStr(strLT, "DAY")
                                            'm8 = InStr(strLT, "WEEK")
                                            'If m8 = 0 Then m8 = InStr(strLT, "DAY")
                                            'If m8 = 0 Then m8 = InStr(strLT, "DYA")
                                            'm9 = InStr(strLT, ",")
                                            'If m9 > 0 And m8 > 0 Then
                                               'If m9 > m8 Then
                                                  'strLT = Trim(Left(strLT, m9 - 1))
                                               'Else
                                                  'strLT = Trim(Right(strLT, Len(strLT) - m9))
                                               'End If
                                            'End If
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            If m6 = 0 Then m6 = InStr(strLT, "～")
                                            m7 = InStr(strLT, "WORKINGDAY")
                                            If m7 = 0 Then m7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WORKINGWEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "WEEK")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 7 + i1 * 3).Value = Replace(Price, ",", "")
                                               .Cells(r1 + 1, 8 + i1 * 3).Value = LT
                                            End If
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            If Price <> "" Then
                                               .Cells(r1 + 1, 10 + i1 * 3).Value = Replace(Price, ",", "")
                                               .Cells(r1 + 1, 11 + i1 * 3).Value = LT
                                            End If
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\WUXI_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("WUXI Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from WUXI have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub


Public Sub IQ_ANGENE()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If (InStr(sAddress, "@angenechemical.com") + InStr(sAddress, "@an-gene.com")) > 0 And (InStr(sTitle, "RE: Inquiry from AstaTech") = 1 _
                  Or InStr(sTitle, "Re: Inquiry from AstaTech") = 1) Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If
                    
                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "can't give help") + InStr(MailText, "不提供") + InStr(MailText, "不报价") _
                    + InStr(MailText, "unavailable") + InStr(MailText, "can't provide") + InStr(MailText, "not available") _
                    + InStr(MailText, "can't supply") + InStr(MailText, "not able to supply") + InStr(MailText, "do not supply") > 0 And _
                    (InStr(MailText, "find our quotation on your requested") + InStr(MailText, "check our quotation as follows")) = 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "ANGENE"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "find our quotation on your requested") + InStr(MailText, "check our quotation as follows") > 0 Then
                          pCAT = ""
                          CAS = ""
                          Desc = ""
                          Purity = ""
                          LT = ""
                          k = 0
                          n = 0
                          i0 = 0
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          If InStr(sAddress, "@angenechemical.com") > 0 Then Set objList = html.getElementsbyTagName("li")
                          If InStr(sAddress, "@an-gene.com") > 0 Then Set objList = html.getElementsbyTagName("div")
                          
                          
                          For j = 0 To objRows.Length
                             On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                             On Error GoTo 0
                              If InStr(strObj, "AST Cat") = 1 Then
                                On Error Resume Next
                                 CAT = Trim(objRows(j).Cells(1).innerText)
                                 MDL = Trim(objRows(j).Cells(3).innerText)
                                On Error GoTo 0
                                 Exit For
                              End If
                          Next j
                          
                          For i = 0 To objList.Length
                              
                             On Error Resume Next
                              strList = objList(i).innerText
                             On Error GoTo 0
                             If n > 0 Then
                                GoTo RepeatItem
                             End If
                             
                             If InStr(strList, "Product ID:") > 0 Then
                                pCAT = Trim(Right(strList, Len(strList) - InStr(strList, ":")))
                             End If
                             If InStr(strList, "Compound Name:") + InStr(strList, "Compound name:") > 0 Then
                                Desc = Trim(Right(strList, Len(strList) - InStr(strList, ":")))
                             End If
                             If InStr(strList, "CAS:") > 0 Then
                                CAS = Trim(Right(strList, Len(strList) - InStr(strList, ":")))
                             End If
                             If InStr(strList, "Purity:") > 0 Then
                                Purity = Trim(Right(strList, Len(strList) - InStr(strList, ":")))
                             End If
                             
                           With WS1
                             'r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                             If Purity <> "" Then
                               'If pCAT <> "" Then
                                'If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                   'GoTo NextI
                                'Else
                                   'n = i
                                   'r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                   '.Cells(r1 + 1, 1).Value = "ANGENE"
                                   '.Cells(r1 + 1, 2).Value = pCAT
                                   '.Cells(r1 + 1, 3).Value = CAT
                                   '.Cells(r1 + 1, 4).Value = rcvdDate
                                   '.Cells(r1 + 1, 5).Value = sentDate
                                   '.Cells(r1 + 1, 6).Value = Sender
                                   '.Cells(r1 + 1, 7).Value = CAS
                                   '.Cells(r1 + 1, 8).Value = MDL
                                   '.Cells(r1 + 1, 9).Value = Purity
                                   '.Cells(r1 + 1, 10).Value = Desc
                                   '.Cells(r1 + 1, 11).Value = Rqmts
                                   '.Cells(r1 + 1, 27).Value = Cmnts
                                   'objMail.UnRead = False
                                   'm = m + 1
                                'End If
                               'Else
                                n = i
                                r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                .Cells(r1 + 1, 1).Value = "ANGENE"
                                .Cells(r1 + 1, 2).Value = pCAT
                                .Cells(r1 + 1, 3).Value = CAT
                                .Cells(r1 + 1, 4).Value = rcvdDate
                                .Cells(r1 + 1, 5).Value = sentDate
                                .Cells(r1 + 1, 6).Value = Sender
                                .Cells(r1 + 1, 7).Value = CAS
                                .Cells(r1 + 1, 8).Value = MDL
                                .Cells(r1 + 1, 9).Value = Purity
                                .Cells(r1 + 1, 10).Value = Desc
                                '.Cells(r1 + 1, 11).Value = Rqmts
                                '.Cells(r1 + 1, 27).Value = Cmnts
                                objMail.UnRead = False
                                m = m + 1
                                'If Font = "#ff0000" Then
                                '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                'End If
                               'End If
                             End If
RepeatItem:
                             If InStr(strList, "Price (EXW):") + InStr(strList, "Price(EXW):") > 0 Then
                                i0 = i
                                For i1 = i0 To objList.Length
                                   On Error Resume Next
                                    strList1 = objList(i1).innerText
                                   On Error GoTo 0
                                    If InStr(strList1, "Lead time:") > 0 Then
                                       If Purity = "" Then GoTo NextI
                                       strLT = Trim(UCase(Replace(Right(strList1, Len(strList1) - InStr(strList1, ":")), " ", "")))
                                        m5 = InStr(strLT, "STOCK")
                                        m6 = InStr(strLT, "-")
                                        If m6 = 0 Then m6 = InStr(strLT, "～")
                                        m7 = InStr(strLT, "WORKINGDAY")
                                        If m7 = 0 Then m7 = InStr(strLT, "DAY")
                                        m8 = InStr(strLT, "WORKINGWEEK")
                                        If m8 = 0 Then m8 = InStr(strLT, "WEEK")

                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     If InStr(UCase(strLT), "FOR") > 0 Then GoTo errorLT
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
                                       Exit For
                                    End If
                                Next i1
                                
                                Do While InStr(strList, "Price") > 0
                                   Qty = ""
                                   Price = ""
                                   strPrice = Trim(Replace(Replace(UCase(Replace(Right(strList, Len(strList) - InStr(strList, ":")), " ", "")), "$", ""), "USD", ""))
                                   
                                   If InStr(strPrice, ";") > 0 Then
                                      Do
                                        If InStr(strPrice, ";") > 0 Then
                                           strPrice0 = Left(strPrice, InStr(strPrice, ";") - 1)
                                        Else
                                           strPrice0 = strPrice
                                        End If
                                        strPrice = Trim(Right(strPrice, Len(strPrice) - InStr(strPrice, ";")))
                                        If InStr(strPrice0, "/") > 0 Then
                                           Price = Left(strPrice0, InStr(strPrice0, "/") - 1)
                                           Qty = Right(strPrice0, Len(strPrice0) - InStr(strPrice0, "/"))
                                           If (12 + 3 * k) < 25 Then
                                              WS1.Cells(r1 + 1, 12 + 3 * k) = Qty
                                              WS1.Cells(r1 + 1, 13 + 3 * k) = Price
                                              WS1.Cells(r1 + 1, 14 + 3 * k) = LT
                                           Else
                                              GoTo NextI
                                           End If
                                           k = k + 1
                                        End If
                                      Loop Until strPrice0 = strPrice
                                   
                                   Else
                                      If InStr(strPrice, "/") > 0 Then
                                          Price = Left(strPrice, InStr(strPrice, "/") - 1)
                                          Qty = Right(strPrice, Len(strPrice) - InStr(strPrice, "/"))
                                          If (12 + 3 * k) < 25 Then
                                              WS1.Cells(r1 + 1, 12 + 3 * k) = Qty
                                              WS1.Cells(r1 + 1, 13 + 3 * k) = Price
                                              WS1.Cells(r1 + 1, 14 + 3 * k) = LT
                                          Else
                                              GoTo NextI
                                          End If
                                          k = k + 1
                                      End If
                                   End If
                                   'k = k + 1
                                   i = i + 1
                                  On Error Resume Next
                                   strList = objList(i).innerText
                                  On Error GoTo 0
                                Loop
                             End If
                             
                             If InStr(strList, "Structure:") + InStr(strList, "Best regards") > 0 Then
                                GoTo NextI
                             End If
                             
                           End With
                          Next i
                          
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\ANGENE_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("ANGENE Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from ANGENE have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub


Public Sub IQ_HUATENG()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@chemreagents.com") > 0 And InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If

                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "提供不了") + InStr(MailText, "不提供") + InStr(MailText, "out of stock") _
                    + InStr(MailText, "not available") + InStr(MailText, "unavailable") + InStr(MailText, "no stock") _
                    + InStr(MailText, "don't have stock") > 0 And (InStr(MailText, "our best quotation as") + InStr(MailText, "make custom synthesis")) = 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "HUATENG"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       'If InStr(MailText, "Partner Cat#") < 100 Then
                       If InStr(MailText, "our best quotation as") + InStr(MailText, "make custom synthesis") + InStr(MailText, "quotation as follows") > 0 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0

                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  'strCmnts = objRows(j2 + 3).Cells(0).innerText
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "HUATENG"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = Replace(Replace(Replace(UCase(Trim(Replace(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", ""), "$", ""))), "USD", ""), "N/A", ""), "NA", "")
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               'm2 = InStr(strPrice, "BY")
                                               'm3 = InStr(strPrice, "(")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               If m1 > 0 Then
                                                  Price = Trim(Left(strPrice, m1 - 1))
                                                  Qty0 = Trim(Right(strPrice, Len(strPrice) - m1))
                                               End If
                                            End If
                                            
                                            If Qty = "" Then Qty = Qty0
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            If m6 = 0 Then m6 = InStr(strLT, "～")
                                            m7 = InStr(strLT, "WORKINGDAY")
                                            If m7 = 0 Then m7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WORKINGWEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "WEEK")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            .Cells(r1 + 1, 7 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            If Price <> "" Then .Cells(r1 + 1, 8 + i1 * 3).Value = GetNum(LT)
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            If Price <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = GetNum(LT)
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\HUATENG_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("HUATENG Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from HUATENG have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub


Public Sub IQ_COOLPHARM()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 CAT = ""
                 compID = ""
                 If InStr(sAddress, "@coolpharm.com") > 0 And _
                (InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Or InStr(sTitle, "Re:Inquiry from AstaTech") = 1) Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If

                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "not available") + InStr(MailText, "unstable") > 0 _
                      And InStr(MailText, "check our offer as follows") = 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "COOLPHARM"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "check our offer as follows") > 0 Then
                          compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                            If strObj = "Partner Cat#:" Then j1 = j
                            If strObj = "Extra Quoted Quantity" Then
                               j2 = j
                               k = j2 - j1 - 5
                               If k > 3 Then k = 3
                               If k > 0 Then
                                  pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                  CAS = Trim(objRows(j1).Cells(3).innerText)
                                  CAT = Trim(objRows(j1 + 1).Cells(1).innerText)

                                  If CAT = "" Or (compID <> pCAT And compID <> CAS And compID <> CAT) Or IsNumeric(Right(CAT, 1)) = False Then
                                     GoTo NextI
                                  End If

                                  MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                  Purity = Trim(Replace(Replace(Replace(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""), "+", ""), ">", ""), "min", ""))
                                  'Font = ""
                                  'Set objPurity = objRows(j1 + 1).getElementsbyTagName("td")
                                  'Set objFont = objPurity(5).getElementsbyTagName("font")
                                  
                                  'On Error Resume Next
                                   'Font = objFont(0).Color
                                  'On Error GoTo 0

                                  Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                  Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                  'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                  'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                  'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                  'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                  'strCmnts = objRows(j2 + 3).Cells(0).innerText
                                  strCmnts = objRows(j2 + 3).innerText
                                  Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                  If InStr(UCase(Cmnts), "HCL") + InStr(UCase(Cmnts), "HBR") + InStr(UCase(Cmnts), "SALT") > 0 Then
                                     GoTo NextI
                                  End If
                                  
                                  With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "COOLPHARM"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     'If Font = "#ff0000" Then
                                        '.Cells(r1 + 1, 28).Value = "Change the Purity"
                                     'End If
                                     For i1 = 1 To k + 3
                                         Price = ""
                                         UPrice = ""
                                         PriceUnit = ""
                                         strPrice = ""
                                         strLT = ""
                                         LT = ""
                                         If i1 = k + 1 Then
                                            GoTo Skip
                                         End If
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         strPrice = UCase(Trim(xlApp.WorksheetFunction.Clean(Replace(Replace(objRows(j1 + 4 + i1).Cells(1).innerText, " ", ""), "$", ""))))
                                         strLT = Trim(Replace(Replace(UCase(objRows(j1 + 4 + i1).Cells(2).innerText), " ", ""), "ABOUT", ""))
                                         If strPrice <> "" Then
                                            If IsNumeric(strPrice) Then
                                               Price = strPrice
                                            Else
                                               m1 = InStr(strPrice, "/")
                                               m2 = InStr(strPrice, "BY")
                                               m3 = InStr(strPrice, "(")
                                               m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                               If m4 > 0 Then
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                                  End If
                                                  GoTo NextI
                                               End If
                                               If m1 > 0 And m2 = 0 And m3 = 0 Then
                                                  UPrice = Trim(Left(strPrice, m1 - 1))
                                                  PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                               Else
                                                  If m1 > 0 And m2 > 0 And m3 = 0 Then
                                                     UPrice = Trim(Left(strPrice, m1 - 1))
                                                     PriceUnit = Trim(Mid(strPrice, m1 + 1, m2 - m1 - 1))
                                                     'ShipPrice = Trim(Mid(strPrice, m2 + 1, m3 - m2 - 1))
                                                  Else
                                                     Note = .Cells(r1 + 1, 28).Value
                                                     If Note <> "" Then
                                                        .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                                     Else
                                                        .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                                     End If
                                                     GoTo NextI
                                                  End If
                                               End If
                                            End If
                                            
                                            If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                               If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                                  QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                               End If
                                            Else
                                               If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                                  QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                                  If PriceUnit = "KG" Then
                                                     Price = CStr(QtyNum * UPrice)
                                                     'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                                  End If
                                                  If PriceUnit = "G" Then
                                                     Price = CStr(QtyNum * UPrice * 1000)
                                                     'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                                  End If
                                               Else
                                                  Note = .Cells(r1 + 1, 28).Value
                                                  If Note <> "" Then
                                                     .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                                  Else
                                                     .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                                  End If
                                                  Exit For
                                               End If
                                            End If
                                            
                                            m5 = InStr(strLT, "STOCK")
                                            m6 = InStr(strLT, "-")
                                            m7 = InStr(strLT, "DAY")
                                            m8 = InStr(strLT, "WEEK")
                                            If m8 = 0 Then m8 = InStr(strLT, "WEEL")
                                            
                                           On Error GoTo errorLT
                                            If m5 > 0 Then
                                               If m7 > 0 Or m8 > 0 Then
                                                  If m6 = 0 Then
                                                     If m8 > 0 Then
                                                        LT = Trim(Left(strLT, m8 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  Else
                                                     If m8 > 0 Then
                                                        LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                     Else
                                                        If m7 > 0 Then
                                                           If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                           Else
                                                              LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  LT = "INSTK"
                                               End If
                                            Else
                                               If m6 = 0 Then
                                                  If m8 > 0 Then
                                                     LT = Trim(Left(strLT, m8 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               Else
                                                  If m8 > 0 Then
                                                     LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                                  Else
                                                     If m7 > 0 Then
                                                        If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                        Else
                                                           LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                           On Error GoTo 0
    
                                         End If
                                         If i1 > k + 1 Then
                                            .Cells(r1 + 1, 6 + i1 * 3).Value = Qty
                                            .Cells(r1 + 1, 7 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            If Price <> "" Then .Cells(r1 + 1, 8 + i1 * 3).Value = GetNum(LT)
                                         Else
                                            .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                            .Cells(r1 + 1, 10 + i1 * 3).Value = Price
                                            'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                            If Price <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = GetNum(LT)
                                         End If
Skip:
                                     Next i1
                                  End With
                               End If
                             Exit For
                            End If
                          Next j
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\COOLPHARM_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("COOLPHARM Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from COOLPHARM have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub


Public Sub IQ_AMADIS()

 Dim olApp, xlApp As Object
 Dim WB1 As Object
 Dim WS1, WS2 As Object
 Dim i, r1, r2 As Long
 Dim html As Object: Set html = CreateObject("htmlfile")
 Dim objRows, objPurity, objFont As Object
 Dim rcvdDate As Date
 Dim filePath As String
 Dim QtyNum As Double
 Const xlUp As Long = -4162

 Set xlApp = GetObject(, "Excel.Application")
 If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
 End If
 With xlApp
    .Application.EnableEvents = False
    .Application.DisplayAlerts = False
    .Application.ScreenUpdating = False
 
 Set WB1 = .Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", True, False)
 Set WS1 = WB1.Sheets(1)
 Set WS2 = WB1.Sheets("Supplier_List")
 r = WS1.Cells(WS1.Rows.Count, "A").End(xlUp).Row
 If r > 1 Then WS1.Rows("2:" & r).ClearContents
 End With
     
 Set olApp = GetObject(, "Outlook.Application")
 If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
 End If
 Set objNamespace = olApp.GetNamespace("MAPI")
 Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
 m = 0
 If objFolder.UnReadItemCount = 0 Then
    MsgBox ("No Unread Emails to process! Please double check your inbox.")
    GoTo CleanUp
 Else
    If objFolder.DefaultItemType = olMailItem Then
       For Each objItem In objFolder.Items
           If objItem.Class = olMail Then
              Set objMail = objItem
              If objMail.UnRead = True Then
                 sAddress = objMail.SenderEmailAddress
                 sTitle = objMail.Subject
                 rcvdDate = Format(objMail.ReceivedTime, "mm/dd/yyyy")
                 pCAT = ""
                 CAT = ""
                 Desc = ""
                 'compID = ""
                 If InStr(sAddress, "@amadischem.com") > 0 And _
                (InStr(sTitle, "Re: Inquiry from AstaTech") = 1 Or InStr(sTitle, "Re:Inquiry from AstaTech") = 1) Then
                    html.Body.Innerhtml = objMail.HTMLBody    'Set the body of the email equal to the html from outlook email
                    MailText = html.Body.innerText
                    If InStr(MailText, "From:") > 0 And InStr(MailText, "Date:") > 0 Then
                       Set objQuote = html.getElementsbyTagName("div")
                       For q = 0 To objQuote.Length
                          On Error Resume Next
                           strQuote0 = Trim(objQuote(q).innerText)
                           strQuote1 = Trim(objQuote(q + 1).innerText)
                           strQuote2 = Trim(objQuote(q + 2).innerText)
                          On Error GoTo 0
                          
                          If InStr(strQuote0, "From:") = 1 And InStr(strQuote1, "Date:") = 1 And InStr(strQuote2, "To:") = 1 Then
                            On Error Resume Next
                             strSender = objQuote(q).getElementsbyTagName("a")(0).getAttribute("href")
                            On Error GoTo 0
                             
                             Sender = Trim(Right(strSender, Len(strSender) - InStr(strSender, ":")))
                             sentDate = Format(Trim(Right(strQuote1, Len(strQuote1) - InStr(strQuote1, ":"))), "mm/dd/yyyy")
                             Exit For
                          End If
                       Next q
                    Else
                       GoTo NextI
                    End If

                    Set objRows = html.getElementsbyTagName("tr")
                    If InStr(MailText, "not available") + InStr(MailText, "unstable") > 0 _
                      And InStr(MailText, "Here is our quotation") = 0 Then
                       For j = 0 To objRows.Length
                         On Error Resume Next
                           strObj = Trim(objRows(j).Cells(0).innerText)
                         On Error GoTo 0
                           
                           If strObj = "Partner Cat#:" Then j1 = j
                           If strObj = "Extra Quoted Quantity" Then
                              j2 = j
                              k = j2 - j1 - 5
                              If k > 3 Then k = 3
                              If k > 0 Then
                                 pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                 CAS = Trim(objRows(j1).Cells(3).innerText)
                                 CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                 If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                    GoTo NextI
                                 End If
                                 MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                 Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                 Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                 Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                 strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                 strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                 EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                 EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                 strCmnts = objRows(j2 + 3).innerText
                                 Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 
                                 With WS1
                                     r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                                     .Cells(r1 + 1, 1).Value = "AMADIS"
                                     .Cells(r1 + 1, 2).Value = pCAT
                                     .Cells(r1 + 1, 3).Value = CAT
                                     .Cells(r1 + 1, 4).Value = rcvdDate
                                     .Cells(r1 + 1, 5).Value = sentDate
                                     .Cells(r1 + 1, 6).Value = Sender
                                     .Cells(r1 + 1, 7).Value = CAS
                                     .Cells(r1 + 1, 8).Value = MDL
                                     .Cells(r1 + 1, 9).Value = Purity
                                     .Cells(r1 + 1, 10).Value = Desc
                                     .Cells(r1 + 1, 11).Value = Rqmts
                                     .Cells(r1 + 1, 27).Value = Cmnts
                                     objMail.UnRead = False
                                     m = m + 1
                                     For i1 = 1 To k
                                         strQty = Trim(Replace(objRows(j1 + 4 + i1).Cells(0).innerText, " ", ""))
                                         strQty = xlApp.WorksheetFunction.Substitute(strQty, "*", "")
                                         Qty = UCase(Trim(Right(strQty, Len(strQty) - InStrRev(strQty, ":"))))
                                         .Cells(r1 + 1, 9 + i1 * 3).Value = Qty
                                         If Qty <> "" Then .Cells(r1 + 1, 11 + i1 * 3).Value = "N/A"
                                     Next i1
                                     If EXQty1 <> "" Then
                                        .Cells(r1 + 1, 21).Value = EXQty1
                                        .Cells(r1 + 1, 23).Value = "N/A"
                                     End If
                                     If EXQty2 <> "" Then
                                        .Cells(r1 + 1, 24).Value = EXQty2
                                        .Cells(r1 + 1, 26).Value = "N/A"
                                     End If
                                 End With
                              End If
                              Exit For
                           End If
                       Next j
                    Else
                       If InStr(MailText, "Here is our quotation") > 0 Then
                          'compID = Trim(Right(sTitle, Len(sTitle) - InStr(sTitle, "-")))
                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj = Trim(objRows(j).Cells(0).innerText)
                            On Error GoTo 0
                           
                              If strObj = "Partner Cat#:" Then j1 = j
                              If strObj = "Extra Quoted Quantity" Then
                                 j2 = j
                                 k = j2 - j1 - 5
                                 If k > 3 Then k = 3
                                 If k > 0 Then
                                    'pCAT = UCase(Trim(objRows(j1).Cells(1).innerText))
                                    'CAS = Trim(objRows(j1).Cells(3).innerText)
                                    CAT = Trim(objRows(j1 + 1).Cells(1).innerText)
                                    If CAT = "" Or IsNumeric(Right(CAT, 1)) = False Then
                                       GoTo NextI
                                    End If
                                    MDL = Trim(objRows(j1 + 1).Cells(3).innerText)
                                    'Purity = Trim(Replace(objRows(j1 + 1).Cells(5).innerText, " ", ""))
                                    'Desc = Trim(objRows(j1 + 2).Cells(1).innerText)
                                    'Rqmts = Trim(objRows(j1 + 3).Cells(1).innerText)
                                    'strEXQty1 = Trim(Replace(objRows(j2 + 1).Cells(0).innerText, " ", ""))
                                    'strEXQty2 = Trim(Replace(objRows(j2 + 2).Cells(0).innerText, " ", ""))
                                    'EXQty1 = UCase(Trim(Right(strEXQty1, Len(strEXQty1) - InStrRev(strEXQty1, ":"))))
                                    'EXQty2 = UCase(Trim(Right(strEXQty2, Len(strEXQty2) - InStrRev(strEXQty2, ":"))))
                                    'strCmnts = objRows(j2 + 3).innerText
                                    'Cmnts = Trim(Right(strCmnts, Len(strCmnts) - InStr(strCmnts, ":")))
                                 End If
                                 Exit For
                              End If
                          Next j
                          
                        With WS1
                           r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
                           .Cells(r1 + 1, 1).Value = "AMADIS"
                           '.Cells(r1 + 1, 2).Value = pCAT
                           .Cells(r1 + 1, 3).Value = CAT
                           .Cells(r1 + 1, 4).Value = rcvdDate
                           .Cells(r1 + 1, 5).Value = sentDate
                           .Cells(r1 + 1, 6).Value = Sender
                           '.Cells(r1 + 1, 7).Value = CAS
                           .Cells(r1 + 1, 8).Value = MDL
                           '.Cells(r1 + 1, 9).Value = Purity
                           '.Cells(r1 + 1, 10).Value = Desc
                           '.Cells(r1 + 1, 11).Value = Rqmts
                           '.Cells(r1 + 1, 27).Value = Cmnts
                           objMail.UnRead = False
                           m = m + 1
                           'If Font = "#ff0000" Then
                           '.Cells(r1 + 1, 28).Value = "Change the Purity"
                           'End If

                          For j = 0 To objRows.Length
                            On Error Resume Next
                              strObj1 = Trim(objRows(j).Cells(0).innerText)
                              strObj2 = Trim(objRows(j).Cells(1).innerText)
                            On Error GoTo 0
                            
                            If strObj1 = "Quantity" And InStr(strObj2, "Price") = 1 Then
                              On Error Resume Next
                               strDesc = Trim(objRows(j - 1).Cells(0).innerText)
                               If strDesc <> "" And InStr(strDesc, "CAT#") > 0 Then
                                  m9 = InStr(strDesc, "CAS#")
                                  If m9 = 0 Then m9 = InStr(strDesc, "CAS")
                                  m10 = InStr(strDesc, "CAT#")
                                  If m10 = 0 Then m10 = InStr(strDesc, "CAT")
                                  If m9 > 0 Then
                                     Desc = Trim(xlApp.WorksheetFunction.Clean(Replace(Left(strDesc, m9 - 1), ";", "")))
                                     CAS = Trim(Replace(Replace(Mid(strDesc, m9 + 4, m10 - m9 - 4), ";", ""), ":", ""))
                                     pCAT = Trim(Replace(Replace(Right(strDesc, Len(strDesc) - m10 - 3), ";", ""), ":", ""))
                                  Else
                                     Desc = Trim(xlApp.WorksheetFunction.Clean(Replace(Left(strDesc, m10 - 1), ";", "")))
                                     pCAT = Trim(Replace(Replace(Right(strDesc, Len(strDesc) - m10 - 3), ";", ""), ":", ""))
                                  End If
                               End If
                              On Error GoTo 0
                              
                               .Cells(r1 + 1, 2).Value = pCAT
                               .Cells(r1 + 1, 7).Value = CAS
                               .Cells(r1 + 1, 10).Value = Desc
                                 
                               n = 1
                              On Error Resume Next
                               strQty = Trim(xlApp.WorksheetFunction.Clean(Replace(objRows(j + n).Cells(0).innerText, " ", "")))
                               strPrice = UCase(Trim(xlApp.WorksheetFunction.Clean(Replace(Replace(objRows(j + n).Cells(4).innerText, " ", ""), "$", ""))))
                               strLT = Trim(Replace(Replace(UCase(objRows(j + n).Cells(3).innerText), " ", ""), "ABOUT", ""))
                               Qty = UCase(strQty)
                              On Error GoTo 0
                               
                               Do Until InStr(strQty, "Purity") = 1 Or n > 5
                                 
                                 If strPrice <> "" Then
                                    If IsNumeric(strPrice) Then
                                       Price = strPrice
                                    Else
                                       m1 = InStr(strPrice, "/")
                                       m2 = InStr(strPrice, "BY")
                                       m3 = InStr(strPrice, "(")
                                       m4 = InStr(strPrice, "HCL") + InStr(strPrice, "SALT") + InStr(strPrice, "HBR")
                                       If m4 > 0 Then
                                          Note = .Cells(r1 + 1, 28).Value
                                          If Note <> "" Then
                                             .Cells(r1 + 1, 28).Value = Note & ";" & "HCl or Salt"
                                          Else
                                             .Cells(r1 + 1, 28).Value = "HCl or Salt"
                                          End If
                                          GoTo NextI
                                       End If
                                       If m1 > 0 And m2 = 0 And m3 = 0 Then
                                          UPrice = Trim(Left(strPrice, m1 - 1))
                                          PriceUnit = Trim(Right(strPrice, Len(strPrice) - m1))
                                       Else
                                          If m1 > 0 And m2 > 0 And m3 = 0 Then
                                             UPrice = Trim(Left(strPrice, m1 - 1))
                                             PriceUnit = Trim(Mid(strPrice, m1 + 1, m2 - m1 - 1))
                                             'ShipPrice = Trim(Mid(strPrice, m2 + 1, m3 - m2 - 1))
                                          Else
                                             Note = .Cells(r1 + 1, 28).Value
                                             If Note <> "" Then
                                                .Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted Price"
                                             Else
                                                .Cells(r1 + 1, 28).Value = "Unformatted Price"
                                             End If
                                             GoTo NextI
                                          End If
                                       End If
                                    End If
                                    
                                    If Right(Qty, 2) = "KG" Or Right(Qty, 2) = "MG" Then
                                       If IsNumeric(Left(Qty, Len(Qty) - 2)) And Right(Qty, 2) = "KG" Then
                                          QtyNum = Left(Qty, Len(Qty) - 2) * 1
                                          If PriceUnit = "KG" Then
                                             Price = CStr(QtyNum * UPrice)
                                             'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                          End If
                                       End If
                                    Else
                                       If Right(Qty, 1) = "G" And IsNumeric(Left(Qty, Len(Qty) - 1)) Then
                                          'QtyNum = Left(Qty, Len(Qty) - 1) / 1000
                                          'If PriceUnit = "KG" Then
                                             'Price = CStr(QtyNum * UPrice)
                                             'Price = CStr(QtyNum * UPrice + ShipPrice * 1)
                                          'End If
                                          'If PriceUnit = "G" Then
                                             'Price = CStr(QtyNum * UPrice * 1000)
                                             'Price = CStr(QtyNum * UPrice * 1000 + ShipPrice * 1)
                                          'End If
                                       Else
                                          Note = .Cells(r1 + 1, 28).Value
                                          If Note <> "" Then
                                             .Cells(r1 + 1, 28).Value = Note & ";" & "Unusual package unit"
                                          Else
                                             .Cells(r1 + 1, 28).Value = "Unusual package unit"
                                          End If
                                          Exit For
                                       End If
                                    End If
                                    
                                    m5 = InStr(strLT, "STOCK")
                                    m6 = InStr(strLT, "-")
                                    m7 = InStr(strLT, "DAY")
                                    m8 = InStr(strLT, "WEEK")
                                    'If m8 = 0 Then m8 = InStr(strLT, "WEEL")
                                            
                                   On Error GoTo errorLT
                                    If m5 > 0 Then
                                       If m7 > 0 Or m8 > 0 Then
                                          If m6 = 0 Then
                                             If m8 > 0 Then
                                                LT = Trim(Left(strLT, m8 - 1))
                                             Else
                                                If m7 > 0 Then
                                                   If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                      LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                   Else
                                                      LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                   End If
                                                End If
                                             End If
                                          Else
                                             If m8 > 0 Then
                                                LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                             Else
                                                If m7 > 0 Then
                                                   If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                      LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                   Else
                                                      LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                   End If
                                                End If
                                             End If
                                          End If
                                       Else
                                          LT = "INSTK"
                                       End If
                                    Else
                                       If m6 = 0 Then
                                          If m8 > 0 Then
                                             LT = Trim(Left(strLT, m8 - 1))
                                          Else
                                             If m7 > 0 Then
                                                If Trim(Left(strLT, m7 - 1)) Mod 7 = 0 Then
                                                   LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7))
                                                Else
                                                   LT = CStr(Int(Trim(Left(strLT, m7 - 1)) / 7) + 1)
                                                End If
                                             End If
                                          End If
                                       Else
                                          If m8 > 0 Then
                                             LT = Trim(Mid(strLT, m6 + 1, m8 - m6 - 1))
                                          Else
                                             If m7 > 0 Then
                                                If Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) Mod 7 = 0 Then
                                                   LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7))
                                                Else
                                                   LT = CStr(Int(Trim(Mid(strLT, m6 + 1, m7 - m6 - 1)) / 7) + 1)
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                   On Error GoTo 0
                                   
                                 End If

                                 .Cells(r1 + 1, 9 + n * 3).Value = Qty
                                 .Cells(r1 + 1, 10 + n * 3).Value = Price
                                 'If IsNumeric(LT) And LT * 1 > 12 Then LT = "12"
                                 If Price <> "" Then .Cells(r1 + 1, 11 + n * 3).Value = GetNum(LT)
                                 
                                 n = n + 1
                                 
                                On Error Resume Next
                                 strQty = Trim(xlApp.WorksheetFunction.Clean(Replace(objRows(j + n).Cells(0).innerText, " ", "")))
                                 strPrice = UCase(Trim(xlApp.WorksheetFunction.Clean(Replace(Replace(objRows(j + n).Cells(4).innerText, " ", ""), "$", ""))))
                                 strLT = Trim(Replace(Replace(UCase(objRows(j + n).Cells(3).innerText), " ", ""), "ABOUT", ""))
                                 Qty = UCase(strQty)
                                On Error GoTo 0
                                 
                               Loop
                               
                               If InStr(strQty, "Purity") = 1 Then
                                  Purity = Trim(Replace(objRows(j + n).Cells(1).innerText, " ", ""))
                                  .Cells(r1 + 1, 9).Value = Purity
                               End If
                               Exit For
                            End If
                               
                          Next j
                          
                        End With
                       Else
                          GoTo NextI
                       End If
                    End If
                 End If
              End If
           End If
NextI:
       Next
       
       filePath = WB1.Path & "\AMADIS_InquiryInfo.xlsx"
       On Error Resume Next
       Kill filePath
       WB1.Sheets("Supplier_List").Delete
       WB1.SaveAs Filename:=filePath
       On Error GoTo 0
       'WB1.Close
       MsgBox ("AMADIS Inquiry file has been created successfully." & vbNewLine _
               & m & " messages from AMADIS have been processed.")
    End If
 End If
 GoTo CleanUp
 
errorLT:
       Note = WS1.Cells(r1 + 1, 28).Value
       If Note <> "" Then
          WS1.Cells(r1 + 1, 28).Value = Note & ";" & "Unformatted LeadTime"
       Else
          WS1.Cells(r1 + 1, 28).Value = "Unformatted LeadTime"
       End If
      On Error GoTo 0
       
       GoTo NextI

CleanUp:
    WB1.Close
    Set WB1 = Nothing
    Set WS1 = Nothing
    Set WS2 = Nothing
    
    With xlApp
       .Application.EnableEvents = True
       .Application.DisplayAlerts = True
       .Application.ScreenUpdating = True
    End With
    
    Set xlApp = Nothing
    Set olApp = Nothing
    Set html = Nothing

End Sub



