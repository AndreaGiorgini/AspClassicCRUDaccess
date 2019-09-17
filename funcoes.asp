<%
'**************************************************************************************************
'                               ARQUIVOS DE FUN��ES ASP DIVERSAS
'**************************************************************************************************
'VARI�VEIS STRINGS
'**************************************************************************************************
'Verifica se o valor passado � vazio, empty ou null retornando true/false
'//=========================================================================
function FblnVazio( strValor )
    FblnVazio = CBool( IsEmpty(strValor) or IsNull(strValor) or Trim(strValor) = "" )
end function
'//=========================================================================
'Retorna com um novo valor caso seja vazio, empty ou null
'//=========================================================================
function FVazio( strValor, strNovoValor )
    Dim strValorFVazio
    if FblnVazio( strValor ) Then
        strValorFVazio = CStr(strNovoValor)
    else
        strValorFVazio = CStr(strValor)
    end if
    FVazio = strValorFVazio
end function
'//=========================================================================
'Condi��o Tern�ria
'//=========================================================================
function IIF( condicao, condicao1, condicao2 )
    if condicao then
        IIF = condicao1
    else
        IIF = condicao2
    end if
end function
'//=========================================================================
'Abrevia palavras com reticencias
'//=========================================================================
function abrevia( strPalavra, tamanho )
    if not FblnVazio(strPalavra) then
        if Len(strPalavra) > tamanho then
            abrevia = "<span title='" & strPalavra & "'>" & Mid( strPalavra, 1, tamanho-3 ) & "...</span>"
        else
            abrevia = strPalavra
        end if
    else
        abrevia = strPalavra
    end if
end function
'//=========================================================================
' Coloca Zeros a esquerda 0000001
'//=========================================================================
Function zerosEsquerda(strString, IntTamString)
  Dim NewStr
  if isNull(strString) then
    Newstr = String(IntTamString,"0")
  else
    NewStr = String(IntTamString - Len(strString),"0") & strString
  End if
  zerosEsquerda = NewStr
End Function
'//=========================================================================
'Executa o valor de uma String (Tipo Eval())
'//=========================================================================
Function ExecuteString( str )
    Execute( "ExecuteString" & " = " & str )
End Function
'//=========================================================================
'Mostra parte do texto na pagina inicial do blog
'//=========================================================================
Function Anteprima(sText, nParole)
  Dim nTemp, nVolte
  sText = Replace(sText, vbCrLf, "")
  nTemp = InStr(sText, " ")
  If nTemp <> 0 Then
     nVolte = 1
     While nTemp <> 0 And nVolte < nParole 
         nVolte = nVolte + 1
         nTemp = InStr(nTemp + 1, sText, " ")
     Wend
  End If
  If nVolte > 0 Then
     If nTemp > 0 Then
         Anteprima = Mid(sText, 1, nTemp - 1) & " ..."
     Else
         Anteprima = sText
     End If
  Else
     If Len(sText) > 0 Then
         Anteprima = sText
     Else
         Anteprima = "" 
     End If
  End If
End Function
'**************************************************************************************************
'VARIAVEIS NUMERICAS
'**************************************************************************************************
'Converte para valor monet�rio
'//=========================================================================
function convValor( val )
    If len(val < 5) and Not FblnVazio(val) then 
       valor = FormatNumber( Replace( val, ".", ","), 2 )
    Else
       valor = val
    End if
    convValor = valor
end function
'//=========================================================================
'Fun��o que converte um n�mero no formato string para double
'//=========================================================================
Function to_double(strNumber)
   Dim strAux, strNum
   strAux = mid(cstr(cdbl(1/2)),2,1)
   strNum = Replace(strNumber, ",", strAux)
   strNum = Replace(strNumber, ".", strAux)
   To_Double = cdbl(strNum)
End Function
'//=========================================================================
'string SomenteNumeros(string Str)
'objetivo: retirar caracteres diferentes de 0..9
'//=========================================================================
Function SomenteNumeros(Str)
 Dim strNum
 StrNum = ""
 For i = 1 To Len(Str)
   If IsNumeric(Mid(Str, i, 1)) Then StrNum = StrNum & Mid(Str, i, 1)
 Next
 SomenteNumeros=StrNum
End Function
'//=========================================================================
'DEFINI��O DE UM NUMERO ALEAT�RIO RAND�NICO
'//=========================================================================
Function RandomNumber(intHighestNumber)	
	Randomize()
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function
'**************************************************************************************************
'Formata CNPJ / CPF
'**************************************************************************************************
'Formata CNPJ XX.XXX.XXX/XXXX-XX ou CPF XXX.XXX.XXX-XX
'//=========================================================================
Function formataCPFCNPJ( strCPFCNPJ )

    If Not FblnVazio(strCPFCNPJ) Then
        strCPFCNPJ = Replace(Replace(Replace(Trim(strCPFCNPJ),".",""),"/",""),"-","")

        If Len( strCPFCNPJ ) = 14 Then
            formataCPFCNPJ = Mid(strCPFCNPJ,1,2) & "." & Mid(strCPFCNPJ,3,3) & "." & Mid(strCPFCNPJ,6,3) & "/" & Mid(strCPFCNPJ,9,4) & "-" & Mid(strCPFCNPJ,12,2) 
        ElseIf Len( strCPFCNPJ ) = 11 Then
            formataCPFCNPJ = Mid(strCPFCNPJ,1,3) & "." & Mid(strCPFCNPJ,4,3) & "." & Mid(strCPFCNPJ,7,3) & "-" & Mid(strCPFCNPJ,10,2)
        Else
            formataCPFCNPJ = strCPFCNPJ
        end if
		
     Else
        formataCPFCNPJ = strCPFCNPJ
     end if
        
End Function
'//=========================================================================
' Funcao para calcular CPF 
'//=========================================================================
Function ValCPF(strCPF)
	Dim RecebeCPF, Numero(11), soma, resultado1, resultado2, retorno
	'Retirar todos os caracteres que nao sejam 0-9
	RecebeCPF = SomenteNumeros(strCPF)
	if len(RecebeCPF) <> 11 then
		retorno = false 
	elseif RecebeCPF = "00000000000" then
		retorno = false 
	else
		Numero(1) = Cint(Mid(RecebeCPF,1,1))
		Numero(2) = Cint(Mid(RecebeCPF,2,1))
		Numero(3) = Cint(Mid(RecebeCPF,3,1))
		Numero(4) = Cint(Mid(RecebeCPF,4,1))
		Numero(5) = Cint(Mid(RecebeCPF,5,1))
		Numero(6) = CInt(Mid(RecebeCPF,6,1))
		Numero(7) = Cint(Mid(RecebeCPF,7,1))
		Numero(8) = Cint(Mid(RecebeCPF,8,1))
		Numero(9) = Cint(Mid(RecebeCPF,9,1))
		Numero(10) = Cint(Mid(RecebeCPF,10,1))
		Numero(11) = Cint(Mid(RecebeCPF,11,1))
		soma = 10 * Numero(1) + 9 * Numero(2) + 8 * Numero(3) + 7 * Numero(4) + 6 * Numero(5) + 5 * Numero(6) + 4 * Numero(7) + 3 * Numero(8) + 2 * Numero(9)
		soma = soma -(11 * (int(soma / 11)))
		if soma = 0 or soma = 1 then
			resultado1 = 0
		else
			resultado1 = 11 - soma
		end if
		if resultado1 = Numero(10) then
			soma = Numero(1) * 11 + Numero(2) * 10 + Numero(3) * 9 + Numero(4) * 8 + Numero(5) * 7 + Numero(6) * 6 + Numero(7) * 5 + Numero(8) * 4 + Numero(9) * 3 + Numero(10) * 2
			soma = soma -(11 * (int(soma / 11)))
			if soma = 0 or soma = 1 then
				resultado2 = 0
			else
				resultado2 = 11 - soma
			end if
			if resultado2 = Numero(11) then
				retorno = true 
			else
				retorno = false
			end if
		else 
			retorno = false
		end if
	end if
	valCPF = retorno
end function
'//=========================================================================
' Funcao para calcular CNPJ 
'//=========================================================================
Function ValCNPJ(strCNPJ)
 Dim RecebeCNPJ, Numero(14), soma, resultado1, resultado2,retorno
 RecebeCNPJ = SomenteNumeros(strCNPJ)
 if len(RecebeCNPJ) <> 14 then
   retorno = false
 elseif RecebeCNPJ = "00000000000000" then
    retorno = false
 else
   Numero(1) = Cint(Mid(RecebeCNPJ,1,1))
   Numero(2) = Cint(Mid(RecebeCNPJ,2,1))
   Numero(3) = Cint(Mid(RecebeCNPJ,3,1))
   Numero(4) = Cint(Mid(RecebeCNPJ,4,1))
   Numero(5) = Cint(Mid(RecebeCNPJ,5,1))
   Numero(6) = CInt(Mid(RecebeCNPJ,6,1))
   Numero(7) = Cint(Mid(RecebeCNPJ,7,1))
   Numero(8) = Cint(Mid(RecebeCNPJ,8,1))
   Numero(9) = Cint(Mid(RecebeCNPJ,9,1))
   Numero(10) = Cint(Mid(RecebeCNPJ,10,1))
   Numero(11) = Cint(Mid(RecebeCNPJ,11,1))
   Numero(12) = Cint(Mid(RecebeCNPJ,12,1))
   Numero(13) = Cint(Mid(RecebeCNPJ,13,1))
   Numero(14) = Cint(Mid(RecebeCNPJ,14,1))
   soma = Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2
   soma = soma -(11 * (int(soma / 11)))
   if soma = 0 or soma = 1 then
     resultado1 = 0
   else
     resultado1 = 11 - soma
   end if
   if resultado1 = Numero(13) then
     soma = Numero(1) * 6 + Numero(2) * 5 + Numero(3) * 4 + Numero(4) * 3 + Numero(5) * 2 + Numero(6) * 9 + Numero(7) * 8 + Numero(8) * 7 + Numero(9) * 6 + Numero(10) * 5 + Numero(11) * 4 + Numero(12) * 3 + Numero(13) * 2
     soma = soma - (11 * (int(soma/11)))
     if soma = 0 or soma = 1 then
       resultado2 = 0
     else
       resultado2 = 11 - soma
     end if
     if resultado2 = Numero(14) then
       retorno = true
     else
       retorno = false
     end if
   else
     retorno = false
   end if
 end if
 ValCNPJ = retorno
end function
'**************************************************************************************************
' CRIPTOGRAFIA E DESCRIPTOGRAFIA
'**************************************************************************************************
' funcao Criptografa a senha via ASCII
'//=========================================================================
Function StringToAscii(str)
	Dim result, x
	StringToAscii = ""
	If Len(str)=0 Then Exit Function
	If Len(str)=1 Then
		result = Asc(Mid(str, 1, 1))
		StringToAscii = Left("000", 3-Len(CStr(result))) & CStr(result)
		Exit Function
	End If
	result = ""
	For x=1 To Len(str)
		result = result & StringToAscii(Mid(str, x, 1))
	Next
	StringToAscii = result
End Function
'//=========================================================================
' funcao desCriptografa a senha via ASCII
'//=========================================================================
Function AsciiToString(str)
	Dim result, x
	AsciiToString = ""
	If Len(str)<3 Then Exit Function
	If Len(str)=3 Then
		AsciiToString = Chr(CInt(str))
		Exit Function
	End If
	result = ""
	For x=1 To Len(str) Step 3
		result = result & AsciiToString(Mid(str, x, 3))
	Next
	AsciiToString = result
End Function
'**************************************************************************************************
'ENCODE E DECODE
'**************************************************************************************************
'IsValidUTF8
'//=========================================================================
function IsValidUTF8(s)
  dim i
  dim c
  dim n
  IsValidUTF8 = false
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      select case n
      case 1
        exit function
      case 2
        if (c and &HE0) <> &HC0 then
          exit function
        end if
      case 3
        if (c and &HF0) <> &HE0 then
          exit function
        end if
      case 4
        if (c and &HF8) <> &HF0 then
          exit function
        end if
      case else
        exit function
      end select
      i = i + n
    else
      i = i + 1
    end if
  loop
  IsValidUTF8 = true 
end function
'//=========================================================================
'DecodeUTF8
'//=========================================================================
function DecodeUTF8(s)
  dim i
  dim c
  dim n
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191 
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s 
end function
'//=========================================================================
'EncodeUTF8
'//=========================================================================
function EncodeUTF8(s)
  dim i
  dim c
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c >= &H80 then
      s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
      i = i + 1
    end if
    i = i + 1
  loop
  EncodeUTF8 = s 
end function
'//=========================================================================
'Codifica caracteres vindo do Server.HTMLEncode
'//=========================================================================
Function HTMLEncode(strEncode)
    Dim strAux

    If IsNull(strEncode) Then
        strAux = strEncode
    Else
        strAux = Replace(strEncode, "'", "\'")
        strAux = Replace(strAux, chr(34), "\'")
    End If	  

    HTMLEncode = strAux
End Function
'//=========================================================================
'Decodifica caracteres vindo do Server.HTMLEncode
'//=========================================================================
Function HTMLDecode(sText)
    Dim I
    sText = Replace(sText, "&quot;", Chr(34))
    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
    sText = Replace(sText, "&nbsp;", Chr(32))
    For I = 1 to 255
        sText = Replace(sText, "&#" & I & ";", Chr(I))
    Next
    HTMLDecode = sText
End Function
'//=========================================================================
'Decodifica caracteres vindo do Server.URLEncode
'//=========================================================================
Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
    URLDecode = sOutput
End Function
'//=========================================================================
'Remove/traduz todas as tags malignas
'//=========================================================================
Private Function removemaligno(ByVal strInput)
	strInput = Replace(strInput, "&", "&amp;"    , 1, -1, 1)
	strInput = Replace(strInput, "#", "&#035;"   , 1, -1, 1)
	strInput = Replace(strInput, "%", "&#037;"   , 1, -1, 1)
	strInput = Replace(strInput, "*", "&#042;"   , 1, -1, 1)
	strInput = Replace(strInput, "\", "&#092;"   , 1, -1, 1)
	strInput = Replace(strInput, "'", "&#146;"   , 1, -1, 1)
	strInput = Replace(strInput, "<", "&lt;"     , 1, -1, 1)
	strInput = Replace(strInput, ">", "&gt;"     , 1, -1, 1)
	'strInput = replace(strInput, "�", "&THORN;"  , 1, -1, 1)
	'strInput = replace(strInput, "�", "&thorn;"  , 1, -1, 1)
	'strInput = replace(strInput, "�", "&szlig;"  , 1, -1, 1)	
	removemaligno = strInput
End Function
'//=========================================================================
Private Function traduzMaligno(ByVal strInput)
	strInput = Replace(strInput, "&amp;"    , "&", 1, -1, 1)
	strInput = Replace(strInput, "&#035;"   , "#", 1, -1, 1)
	strInput = Replace(strInput, "&#037;"   , "%", 1, -1, 1)
	strInput = Replace(strInput, "&#042;"   , "*", 1, -1, 1)
	strInput = Replace(strInput, "&#092;"   , "\", 1, -1, 1)
	strInput = Replace(strInput, "&#146;"   , "'", 1, -1, 1)
	strInput = Replace(strInput, "&lt;"     , "<", 1, -1, 1)
	strInput = Replace(strInput, "&gt;"     , ">", 1, -1, 1)
	'strInput = replace(strInput, "&THORN;"  , "�", 1, -1, 1)
	'strInput = replace(strInput, "&thorn;"  , "�", 1, -1, 1)
	'strInput = replace(strInput, "&szlig;"  , "�", 1, -1, 1)	
	traduzMaligno = strInput
End Function
'//=========================================================================
'Remove/traduz todas as tags malignas
'//=========================================================================
'Private Function removemaligno(ByVal strInput)
'	strInput = Replace(strInput, "&", "&amp;"    , 1, -1, 1)
'	strInput = Replace(strInput, "#", "&#035;"   , 1, -1, 1)
'	strInput = Replace(strInput, "%", "&#037;"   , 1, -1, 1)
'	strInput = Replace(strInput, "*", "&#042;"   , 1, -1, 1)
'	strInput = Replace(strInput, "\", "&#092;"   , 1, -1, 1)
'	strInput = Replace(strInput, "'", "&#146;"   , 1, -1, 1)
'	strInput = Replace(strInput, "<", "&lt;"     , 1, -1, 1)
'	strInput = Replace(strInput, ">", "&gt;"     , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Aacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&aacute;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Acirc;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&acirc;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Agrave;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&agrave;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Aring;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&aring;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Atilde;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&atilde;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Auml;"   , 1, -1, 1)
'	strInput = replace(strInput, "�", "&auml;"   , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&AElig;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&aelig;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Eacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&eacute;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ecirc;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ecirc;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Egrave;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&egrave;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Euml;"   , 1, -1, 1)
'	strInput = replace(strInput, "�", "&euml;"   , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&ETH;"    , 1, -1, 1)
'	strInput = replace(strInput, "�", "&eth;"    , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Iacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&iacute;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Icirc;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&icirc;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Igrave;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&igrave;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Iuml;"   , 1, -1, 1)
'	strInput = replace(strInput, "�", "&iuml;"   , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Oacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&oacute;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ocirc;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ocirc;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ograve;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ograve;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Oslash;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&oslash;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Otilde;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&otilde;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ouml;"   , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ouml;"   , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Uacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&uacute;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ucirc;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ucirc;"  , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ugrave;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ugrave;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Uuml;"   , 1, -1, 1)
'	strInput = replace(strInput, "�", "&uuml;"   , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ccedil;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ccedil;" , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Ntilde;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&ntilde;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&reg;"    , 1, -1, 1)
'	strInput = replace(strInput, "�", "&copy;"   , 1, -1, 1)
'	'strInput = replace(strInput, "�", "&Yacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&yacute;" , 1, -1, 1)
'	strInput = replace(strInput, "�", "&THORN;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&thorn;"  , 1, -1, 1)
'	strInput = replace(strInput, "�", "&szlig;"  , 1, -1, 1)	
'	removemaligno = strInput
'End Function
''//=========================================================================
'Private Function traduzMaligno(ByVal strInput)
'	strInput = Replace(strInput, "&amp;"    , "&", 1, -1, 1)
'	strInput = Replace(strInput, "&#035;"   , "#", 1, -1, 1)
'	strInput = Replace(strInput, "&#037;"   , "%", 1, -1, 1)
'	strInput = Replace(strInput, "&#042;"   , "*", 1, -1, 1)
'	strInput = Replace(strInput, "&#092;"   , "\", 1, -1, 1)
'	strInput = Replace(strInput, "&#146;"   , "'", 1, -1, 1)
'	strInput = Replace(strInput, "&lt;"     , "<", 1, -1, 1)
'	strInput = Replace(strInput, "&gt;"     , ">", 1, -1, 1)
'	strInput = replace(strInput, "&Aacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&aacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Acirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&acirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Agrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&agrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Aring;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&aring;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Atilde;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&atilde;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Auml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&auml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&AElig;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&aelig;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Eacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&eacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ecirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ecirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Egrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&egrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Euml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&euml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ETH;"    , "�", 1, -1, 1)
'	strInput = replace(strInput, "&eth;"    , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Iacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&iacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Icirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&icirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Igrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&igrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Iuml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&iuml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Oacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&oacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ocirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ocirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ograve;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ograve;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Oslash;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&oslash;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Otilde;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&otilde;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ouml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ouml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Uacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&uacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ucirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ucirc;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ugrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ugrave;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Uuml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&uuml;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ccedil;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ccedil;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Ntilde;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&ntilde;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&reg;"    , "�", 1, -1, 1)
'	strInput = replace(strInput, "&copy;"   , "�", 1, -1, 1)
'	strInput = replace(strInput, "&Yacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&yacute;" , "�", 1, -1, 1)
'	strInput = replace(strInput, "&THORN;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&thorn;"  , "�", 1, -1, 1)
'	strInput = replace(strInput, "&szlig;"  , "�", 1, -1, 1)	
'	traduzMaligno = strInput
'End Function
'**************************************************************************************************
'ARQUIVOS DIVERSOS
'Parametro deve vir obrigatoriamente do objeto FileSystemObject (file.Size)
'**************************************************************************************************
' Formats given size in bytes,KB,MB and GB
' Set FSO  = server.CreateObject ("Scripting.FileSystemObject")
' Set file = FSO.GetFile (filePath)
' tamanho  = FormatSize(file.Size)
' tipo     = file.Type
'**************************************************************************************************
Function FormatSize(givenSize)
	If (givenSize < 1024) Then
		FormatSize = givenSize & " B"
	ElseIf (givenSize < 1024*1024) Then
		FormatSize = FormatNumber(givenSize/1024,2) & " KB"
	ElseIf (givenSize < 1024*1024*1024) Then
		FormatSize = FormatNumber(givenSize/(1024*1024),2) & " MB"
	Else
		FormatSize = FormatNumber(givenSize/(1024*1024*1024),2) & " GB"
	End If
End Function
'//=========================================================================
' ValidFileExtension()
'//=========================================================================
Function ValidFileExtension(strFileName, strFileExtensions)
    Dim arrExtension
    Dim strFileExtension
    Dim i
	strFileExtension = UCase(GetFileExtension(strFileName))
    arrExtension = Split(UCase(strFileExtensions), ";")
    For i = 0 To UBound(arrExtension)
        'Check to see if a "dot" exists
        If Left(arrExtension(i), 1) = "." Then
            arrExtension(i) = Replace(arrExtension(i), ".", vbNullString)
        End If
        'Check to see if FileExtension is allowed
        If arrExtension(i) = strFileExtension Then
            ValidFileExtension = True
            Exit Function
        End If
    Next
    ValidFileExtension = False
End Function
'//=========================================================================
' InValidFileExtension()
'//=========================================================================
Function InValidFileExtension(strFileName, strFileExtensions)
    Dim arrExtension
    Dim strFileExtension
    Dim i
    strFileExtension = UCase(GetFileExtension(strFileName))
    'Response.Write "filename : " & strFileName & "<br>"
    'Response.Write "file extension : " & strFileExtension & "<br>"    
    'Response.Write strFileExtensions & "<br>"
    'Response.End 
    arrExtension = Split(UCase(strFileExtensions), ";")
    For i = 0 To UBound(arrExtension)
        'Check to see if a "dot" exists
        If Left(arrExtension(i), 1) = "." Then
            arrExtension(i) = Replace(arrExtension(i), ".", vbNullString)
        End If
        
        'Check to see if FileExtension is not allowed
        If arrExtension(i) = strFileExtension Then
            InValidFileExtension = False
            Exit Function
        End If
    Next
    InValidFileExtension = True
End Function
'//=========================================================================
' GetFileExtension()
'//=========================================================================
Function GetFileExtension(strFileName)
    GetFileExtension = Mid(strFileName, InStrRev(strFileName, ".") + 1)
End Function
'**************************************************************************************************
'DATAS
'**************************************************************************************************
'Fun��o para inverter a string de data de ddmmyyyy para yyyymmdd
'//=========================================================================
Function invertdata(strData)
   strData = Replace(strData,"/","")
   a = mid(strData,5,4)
   b = mid(strData,3,2)
   c = mid(strData,1,2)
   strData = a&b&c
   invertdata = (strData)
End Function
'//=========================================================================
'Retorna Data de hoje formatada DD/MM/YYYY
'//=========================================================================
Function getDataBrasil()
    getDataBrasil = Right("00"&Day(Date()),2) & "/" & Right("00"&Month(Date()),2) & "/" & Right("0000"&Year(Date()),4)
End Function
'//=========================================================================
' Retorna o n�mero do dia da semana
'//=========================================================================
Function getNrDiaSemana( strSemana )
    Dim strDia : strDia = "1"
    Dim arrSem(6)
        arrSem(0) = "SAB"
        arrSem(1) = "DOM"
        arrSem(2) = "SEG"
        arrSem(3) = "TER"
        arrSem(4) = "QUA"
        arrSem(5) = "QUI"
        arrSem(6) = "SEX"
    For i=0 To UBound(arrSem)
        If arrSem(i) = UCase(strSemana) Then
            strDia = i
            Exit For
        End If
    Next
    getNrDiaSemana = strDia
End Function
'//=========================================================================
'Fun��o compara as datas se est�o no intervalo ou s�o iguais, retorna True ou False
'Enviar data no formato YYYYMMDD obrigatoriamente
'//=========================================================================
Function DataNoIntervalo( strDataPrincipal, strDataInicio, strDataFinal )
    DataNoIntervalo = CBool( (StrComp(strDataPrincipal,strDataInicio)=1 Or strDataPrincipal=strDataInicio) And (StrComp(strDataPrincipal,strDataFinal)=-1 Or strDataPrincipal=strDataFinal) )
End Function
'//=========================================================================
'Calcula intervalo entre 2 horas enviadas, retornando a hora HH:MM:SS
'//=========================================================================
Function calculaHora( hora1, hora2, blnSinal )
    On Error Resume Next
    'If FblnVazio(hora1) Or Not IsNumeric(Replace(hora1,":","")) Then
    '    hora1 = TimeValue(Time())
    'End If
    'If FblnVazio(hora2) Or Not IsNumeric(Replace(hora2,":","")) Then
    '    hora2 = TimeValue(Time())
    'End If    
    If hora2 = "00:00:00" Then
        horaMomento = DateDiff("s",hora1,TimeValue(Time()))
    Else
        horaMomento = DateDiff("s",hora1,hora2)
    End If
    If Err.number = 13 Then
        horaMomento = DateDiff("s",FVazio(hora1,FVazio(hora2,TimeValue(Time()))),FVazio(hora2,TimeValue(Time())))
        Err.Clear
    End If
    sinal = IIf( horaMomento < 0, "+", "-" )
    Horas = horaMomento \ 3600
    Minutos = (horaMomento mod 3600) \ 60
    Segundos = (horaMomento mod 3600) mod 60
    If blnSinal Then
        calculaHora = sinal & TimeSerial(Horas, Minutos, Segundos)
    Else
        calculaHora = TimeSerial(Horas, Minutos, Segundos)
    End If
End Function
'//=========================================================================
'Formata Data - formataEntrada YYYYMMDD OU DDMMYYYY OU DD/MM/YYYY ou vice-versa
'//=========================================================================
Function FormataData(variavel, formatoEntrada, Forma)
    Dim Dia, Mes, Ano, Ano2D, datData
	datData = variavel	
    Session.LCID = 1046
    If FblnVazio(datData) Then 
        FormataData = datData
        Exit Function
    End If
    datData = Replace(Replace(Replace(datData," ",""),"-",""),"/","")
    select case UCase(formatoEntrada)
        case "YYYYMMDD" :
            datData = Mid(datData,7,2)&"/"&Mid(datData,5,2)&"/"&Mid(datData,1,4)
        case "DDMMYYYY" :
            datData = Mid(datData,1,2)&"/"&Mid(datData,3,2)&"/"&Mid(datData,5,4)
		case "YYYY-MM-AA" : 
			datData = Mid(datData,7,2)&"-"&Mid(datData,5,2)&"-"&Mid(datData,1,4)
        case "DD/MM/AAAA" :
            datData = Mid(datData,1,2)&"/"&Mid(datData,3,2)&"/"&Mid(datData,5,4)			
    end select
    If Not IsDate(datData) Then  
      datData = Now() 
    End If  
    Dia = "" & Right("00" & Cstr(Day(datData)), 2)  
    Mes = "" & Right("00" & Cstr(Month(datData)), 2)  
    Ano = "" & Right("0000" & Cstr(Year(datData)), 4)  
    Ano2D = "" & Right("00" & Cstr(Year(datData)), 2) 
    If Forma = 1 Then  
      FormataData = CStr(Trim(Dia) & "/" & Trim(Mes) & "/" & Trim(Ano))  
    ElseIf Forma = 2 Then  
      FormataData = CStr(Trim(Ano) & "/" & Trim(Mes) & "/" & Trim(Dia)) 
    ElseIf Forma = 3 Then  
      FormataData = CStr(Trim(Ano) & "-" & Trim(Mes) & "-" & Trim(Dia))  
    ElseIf Forma = 4 Then 
      FormataData = CStr(Trim(Dia) & "/" & Trim(Mes) & "/" & Trim(Ano2D)) 
	ElseIf Forma = 5 then
	  FormataData = CStr(Trim(Ano) & "-" & Trim(Mes) & "-" & Trim(DIA)) 
    End If 
End Function
'//=========================================================================
' funcao Formata Data para a configura��o brasileira
'//=========================================================================
Function DataBrasil(strData)
	ano 	  = left(strData, 4)
	mess 	  = right(strData,4)
	mes 	  = left(mess, 2)
	dia 	  = right(strData, 2)
	DataBrasil= dia&"/"&mes&"/"&ano
End Function
'//=========================================================================
Function QuantosDiasTemOMes(Mes,Ano)
  Select Case Mes
    Case 1,3,5,7,8,10,12: QuantosDiasTemOMes = 31
    Case 4,6,9,11: QuantosDiasTemOMes = 30
    Case Else
      If Ano Mod 4 = 0 And (Ano Mod 100 <> 0 Or Ano Mod 400 = 0) Then
        QuantosDiasTemOMes = 29
      Else
        QuantosDiasTemOMes = 28
      End If
  End Select
End Function
'//=========================================================================
Function converteHorarioSegundos(ByVal horario) 
        Dim segundos, minutos, horas 
        Dim horarioArray 
        segundos = "00" 
        minutos  = "00" 
        horas    = "00" 
        horarioArray = split(horario,":") 
        If (uBound(horarioArray) >= 0) Then 
                horas = horarioArray(0) 
        End If 
        If (uBound(horarioArray) >= 1) Then 
                minutos = horarioArray(1) 
        End If 
        If (uBound(horarioArray) >= 2) Then 
                segundos = horarioArray(2) 
        End If 
        converteHorarioSegundos = (horas*3600) + (minutos*60) + segundos 
End Function 
'//=========================================================================
Function converteSegundosHorario(segundos) 
        Dim minutos, horas 
        horas = int(segundos/3600) 
        minutos = int((segundos mod 3600) / 60) 
        segundos = int((segundos mod 3600) mod 60) 
        if len(minutos) < 2 then : minutos = 0 & minutos : end if 
        if len(segundos) < 2 then : segundos = 0 & segundos : end if 
        converteSegundosHorario = horas & ":" & minutos & ":" & segundos 
End Function 
'//=========================================================================
Function SaldoDia(total)
'//Formata converte resultado para horas
'//1.Horario completo do dia
	Horas = total \ 3600
	if isnull(Horas) or (Horas = 0) then horas = "00"end if
	H = len(Horas)
	if H = 1 then Hs = "0"&Horas else Hs = Horas end if
	Minutos = (total mod 3600) \ 60
	if isnull(Minutos) or (Minutos = 0) then Minutos = "00" end if
	if Minutos < 0 then Minutos = (-1 * Minutos) end if
	M = len(Minutos)
	if M = 1 then Ms = "0"&Minutos else Ms = Minutos end if
	Segundos = (total mod 3600) mod 60 
	if isnull(Segundos) or (Segundos = 0) then Segundos = "00" end if
	if Segundos < 0 then Segundos = (-1 * Segundos) end if
	S = len(Segundos)
	if S = 1 then Ss = "0"&Segundos else Ss = Segundos end if
	SaldoDia = Hs&":"&Ms&":"&Ss
End function
'**************************************************************************************************
'IMAGENS
'**************************************************************************************************
' funcao Busca width e height da imagem
'//=========================================================================
function GetBytes(flnm, offset, bytes)
 Dim objFSO
 Dim objFTemp
 Dim objTextStream
 Dim lngSize
 on error resume next
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 Set objFTemp = objFSO.GetFile(flnm)
 lngSize = objFTemp.Size
 set objFTemp = nothing
 fsoForReading = 1
 Set objTextStream = objFSO.OpenTextFile(flnm, fsoForReading)
 if offset > 0 then
    strBuff = objTextStream.Read(offset - 1)
 end if
 if bytes = -1 then        ' Get All!'
    GetBytes = objTextStream.Read(lngSize)  'ReadAll'
 else
    GetBytes = objTextStream.Read(bytes)
 end if
 objTextStream.Close
 set objTextStream = nothing
 set objFSO = nothing
end function

function lngConvert(strTemp)
 lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
end function

function lngConvert2(strTemp)
 lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
end function

function gfxSpex(flnm, width, height, depth, strImageType)
 dim strPNG 
 dim strGIF
 dim strBMP
 dim strType
 strType = ""
 strImageType = "(unknown)"
 gfxSpex = False
 strPNG = chr(137) & chr(80) & chr(78)
 strGIF = "GIF"
 strBMP = chr(66) & chr(77)
 strType = GetBytes(flnm, 0, 3)
 if strType = strGIF then                ' is GIF'
    strImageType = "GIF"
    Width = lngConvert(GetBytes(flnm, 7, 2))
    Height = lngConvert(GetBytes(flnm, 9, 2))
    Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
    gfxSpex = True
 elseif left(strType, 2) = strBMP then        ' is BMP'
    strImageType = "BMP"
    Width = lngConvert(GetBytes(flnm, 19, 2))
    Height = lngConvert(GetBytes(flnm, 23, 2))
    Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
    gfxSpex = True
 elseif strType = strPNG then            ' Is PNG'
    strImageType = "PNG"
    Width = lngConvert2(GetBytes(flnm, 19, 2))
    Height = lngConvert2(GetBytes(flnm, 23, 2))
    Depth = getBytes(flnm, 25, 2)
    select case asc(right(Depth,1))
       case 0
          Depth = 2 ^ (asc(left(Depth, 1)))
          gfxSpex = True
       case 2
          Depth = 2 ^ (asc(left(Depth, 1)) * 3)
          gfxSpex = True
       case 3
          Depth = 2 ^ (asc(left(Depth, 1)))  '8'
          gfxSpex = True
       case 4
          Depth = 2 ^ (asc(left(Depth, 1)) * 2)
          gfxSpex = True
       case 6
          Depth = 2 ^ (asc(left(Depth, 1)) * 4)
          gfxSpex = True
       case else
          Depth = -1
    end select
 else
    strBuff = GetBytes(flnm, 0, -1)        ' Get all bytes from file'
    lngSize = len(strBuff)
    flgFound = 0
    strTarget = chr(255) & chr(216) & chr(255)
    flgFound = instr(strBuff, strTarget)
    if flgFound = 0 then
       exit function
    end if
    strImageType = "JPG"
    lngPos = flgFound + 2
    ExitLoop = false
    do while ExitLoop = False and lngPos < lngSize
       do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
          lngPos = lngPos + 1
       loop
       if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
          lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
          lngPos = lngPos + lngMarkerSize  + 1
       else
          ExitLoop = True
       end if
   loop
   if ExitLoop = False then
      Width = -1
      Height = -1
      Depth = -1
   else
      Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
      Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
      Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
      gfxSpex = True
   end if
 end if
end function
'como chamar
'		Set objFS = Server.CreateObject("Scripting.FileSystemObject")
'		Set objFile = objFS.GetFile("c:\imagem.jpg")
'		If gfxSpex(objFile.Path, w, h, c, strType) = True then
'		  Response.Write " Imagem: <b>" & objFile.name & "</b><br>"
'		  Response.Write "Tamanho: <b>" & w & "x" & h & "</b>"
'		End If
'		Set objFile = Nothing
'		Set objFS = Nothing
'**************************************************************************************************
%>