Attribute VB_Name = "ExportInquiryInfo"
Option Explicit
Sub IQ_ExportInquiryInfo()
'This macro will export partner's inquiry/quote information from unread mails into a excel file.
'Please set up the automatic forward rules in your Outlook before using the Macro for FEIBO.

     Dim xlApp, olApp As Object
     Dim olClass As New GetInquiryInfo
     Dim objNamespace, objFolder As Object
     Dim WB1 As Object
     Dim WS1, WS2 As Object
     Dim Partner, Supplier, ModName, MacName As String
     Dim i, r1, r2 As Long
     Const xlUp As Long = -4162
     
     Set olApp = CreateObject("Outlook.Application")
     Set objNamespace = olApp.GetNamespace("MAPI")
     Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
          
     Set xlApp = CreateObject("Excel.Application")
     With xlApp
          .Application.DisplayAlerts = False
          .Application.ScreenUpdating = False
          .Application.EnableEvents = False
     End With
     
     Set WB1 = xlApp.Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", ReadOnly:=True)
     Set WS1 = WB1.Sheets(1)
     Set WS2 = WB1.Sheets("Supplier_List")
     Partner = InputBox("Please enter the partner's name:" & vbNewLine & "(For example: FDC, INFOARK, PHARMABLOCK, SUNWAY etc.)")
     If Partner = "" Then
        WB1.Close
        GoTo CleanUp
     End If
     With WS2          'WB1.Sheets("Supplier_List")
       Partner = UCase(Trim(Partner))
       If Partner = "ALL" Then
          If objFolder.UnReadItemCount = 0 Then
             WB1.Close
             MsgBox ("No Unread Emails to process! Please double check your inbox.")
             GoTo CleanUp
          Else
             r1 = .Cells(.Rows.Count, "A").End(xlUp).Row
             If r1 > 1 Then
                For i = 2 To r1
                    If i > 2 Then
                       Set WB1 = xlApp.Workbooks.Open("C:\AST_InquiryFiles\AST_PartnerInquiryInfo.xlsx", ReadOnly:=True)
                       Set WS1 = WB1.Sheets(1)
                       Set WS2 = WB1.Sheets("Supplier_List")
                    End If
                    Supplier = WS2.Cells(i, "A").Value
                    MacName = "IQ_" & Supplier
                    CallByName olClass, MacName, VbMethod
                Next i
                MsgBox ("All partner inquiry mails have been processed successfully." & vbNewLine _
                  & "Please check unread mails and process the remaining inquiry mails manually.")
             End If
          End If
       Else
          If xlApp.WorksheetFunction.Countif(.Columns("A"), Partner) > 0 Then
             Supplier = Partner
             MacName = "IQ_" & Supplier
             'Application.Run doesn't work in Outlook
             CallByName olClass, MacName, VbMethod
             'Call IQ_AMADIS
          Else
             MsgBox ("Stopping because the program for " & Partner & " has not been created.")
             WB1.Close
             GoTo CleanUp
          End If
       End If
     End With

CleanUp:
     'WB1.Close Savechanges:=False
     Set WB1 = Nothing
     Set WS1 = Nothing
     Set WS2 = Nothing

     With xlApp
          .Application.DisplayAlerts = True
          .Application.ScreenUpdating = True
          .Application.EnableEvents = True
     End With
     Set xlApp = Nothing
     Set olApp = Nothing
     Set olClass = Nothing

End Sub
                   
            
