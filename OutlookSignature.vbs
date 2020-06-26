On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

GetUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & GetUser)

GetName = objUser.FullName
GetTitle = objUser.Title
GetDepartment = objUser.Department
GetCompany = objUser.Company
GetPhone = objUser.TelephoneNumber
GetOtherPhone = objUser.otherTelephone
GetMobile = objUser.Mobile
GetEmail = objUser.EmailAddress
GetFax = objUser.FaxNumber
GetStreet = objUser.StreetAddress
GetZip = objUser.PostalCode
GetCity = objUser.l
GetState = objUser.State
GetHomepage = objUser.Homepage
GetNotes = objUser.Info
Appendix = "Bei Fragen stehen wir Ihnen gerne zur Verf" & Chr(252) & "gung."
Regards = "Freundliche Gr" & Chr(252) & "sse"

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

' Beginning of signature block

objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 11
objSelection.TypeText Appendix
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.TypeText Regards
objSelection.TypeParagraph()
objSelection.TypeText GetName
objSelection.TypeParagraph()
objSelection.TypeText GetTitle & " " & GetDepartment
objSelection.TypeParagraph()
objSelection.TypeText "Mobile " & GetMobile
objSelection.TypeParagraph()
objSelection.Font.Name = "Arial Black"
objSelection.Font.Size = 13
objSelection.Font.Bold = True
objSelection.Font.Color = RGB(22,46,106)
objSelection.TypeText GetCompany
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 11
objSelection.Font.Bold = False
objSelection.Font.Color = RGB(0,0,0)
objSelection.TypeParagraph()
objSelection.TypeText GetStreet & "  " & GetZip & " " & GetCity & "  " & GetNotes
objSelection.TypeParagraph()
objSelection.TypeText "Telefon " & GetPhone & "  Fax " & GetFax & "  " & GetEmail & "  " & GetHomepage

' End of signature block

Set objSelection = objDoc.Range()

objSignatureEntries.Add "MySignature", objSelection
objSignatureObject.NewMessageSignature = "MySignature"
objSignatureObject.ReplyMessageSignature = "MySignature"

objDoc.Saved = True
objWord.Quit
