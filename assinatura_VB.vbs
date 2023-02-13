On Error Resume Next
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
With objUser
  strName = .FullName
  strTitle = .Description
End With

'Pegar dados do AD, Get data from AD
'Link com todos os dados do AD, Link with all AD data attribute
'https://ss64.com/ps/syntax-ldap.html
strCompany = objUser.Company
strAddress = objUser.streetAddress
strpostalCode = objUser.postalCode
strTitle = objUser.title
strRamal = objUser.homePhone
strPhone = objUser.TelephoneNumber
strMobile = objUser.mobile
strMail = objuser.mail
strWeb = objuser.wWWHomePage
strinstagram = "https://www.instagram.com/"
strlinkedin = "http://twitter.com/"


Set objword = CreateObject("Word.Application")
With objword 
  Set objDoc = .Documents.Add()
  Set objSelection = .Selection
  Set objEmailOptions = .EmailOptions  
  Set objRange = objDoc.Range()
  objDoc.Tables.Add objRange,7,2 'Setando numero de linhas e colunas na tabela, Set number rows and columns in table
  Set objTable = objDoc.Tables(1)
End With
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

With objSelection
	objTable.Columns.Width =  200
	objTable.Cell(1, 1).Merge objTable.Cell(7, 1)	
	
	'Inserir imagem na coluna da esquerda, insert image left column
	objSelection.InlineShapes.AddPicture("\\local-da-logo-na-pasta")	
	
	'nome do usuário, user name
	objTable.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(1, 2).Range.Font.Bold = True
    objTable.Cell(1, 3).Range.Font.Size = "9"
    objTable.Cell(1, 3).Range.Font.Name = "Arial"
	objTable.Cell(1, 2).Range.Font.Color = RGB(105,105,105)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(1, 2).Range.Text = strName 

	'Cargo do usuario, Title user
	objTable.Cell(2, 2).Range.ParagraphFormat.SpaceAfter = 10
	objTable.Cell(2, 2).Range.Font.Size = "9"
    objTable.Cell(2, 2).Range.Font.Name = "Arial"
	objTable.Cell(2, 2).Range.Font.Color = RGB(128,128,128)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(2, 2).Range.Text = strTitle
	
	'Email
	objTable.Cell(3, 2).Range.ParagraphFormat.SpaceAfter = 10
	objTable.Cell(3, 2).Range.Font.Bold = True
	objTable.Cell(3, 2).Range.Font.Size = "9"
    objTable.Cell(3, 2).Range.Font.Name = "Arial"
	objTable.Cell(3, 2).Range.Font.Color = RGB(128,128,128)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)	
	objTable.Cell(3, 2).Range.Text = strMail
	
	'Numero de telefone e ramal, Telephone & extension number 
	objTable.Cell(4, 2).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(4, 2).Range.Font.Size = "9"
    objTable.Cell(4, 2).Range.Font.Name = "Arial"
	objTable.Cell(4, 2).Range.Font.Color = RGB(128,128,128)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(4, 2).Range.Text = "Tel: " & strPhone & "    " & "Ramal: " & strRamal

	'Número do celular, mobile phone
	objTable.Cell(5, 2).Range.ParagraphFormat.SpaceAfter = 10
	objTable.Cell(5, 2).Range.Font.Size = "9"
    objTable.Cell(5, 2).Range.Font.Name = "Arial"
	objTable.Cell(5, 2).Range.Font.Color = RGB(128,128,128)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(5, 2).Range.Text = "Cel: " & strMobile

	'site, website
	objTable.Cell(6, 2).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(6, 2).Range.Font.Bold = True
	objTable.Cell(6, 2).Range.Font.Size = "9"
	objTable.Cell(6, 2).Range.Font.Name = "Arial"
	objTable.Cell(6, 2).Range.Font.Color = RGB(128,128,128)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(6, 2).Range.Text = strweb
	
	'adicionando link e icone de rede social, add hyperlink and social network icon
	objTable.Cell(7, 2).Range.Select.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(7, 2).Range.Select.Text = objDoc.Hyperlinks.Add (objSelection.InlineShapes.AddPicture("\\local_assinatura\instagram.png"), strinstagram)
	objTable.Cell(7, 2).Range.Select.Text = objDoc.Hyperlinks.Add (objSelection.InlineShapes.AddPicture("\\local_assinatura\linkedin.png"), strlinkedin)
	
	.TypeText Chr(1)
	
	'fim da tabela, end table
	.EndKey end_table
	
 
End With

Set objSelection = objDoc.Range()
objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"
objDoc.Saved = True
objword.Quit