REM --------------------------------------------------------------------------------------------------------------------
REM Script: List_Domain_Users.vbs - Lista informações a respeito dos usuários criados no Active Directory
REM Na linha 15 altere domain.name pelo nome do seu domínio e na Linha 19 defina onde o arquivo com os dados dos usuários será salvo. Por padrão eu deixei d:\domain_users.htm
REM Contato: Paulo Roberto Sant´anna Cardoso (contato@paulosantanna.com)
REM Compatibilidade: Windows Server 2008;Windows Server 2012;Windows Server 2016
REM Blog: paulosantanna.com
REM ---------------------------------------------------------------------------------------------------------------------

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const TristateUseDefault = -2
Const TristateTrue = -1
Const TristateFalse = 0
Set prov = getObject("WinNT://domain.name")
dim members

Set fs = CreateObject ("Scripting.FileSystemObject")
Set f = fs.OpenTextFile ("d:\domain_users.htm", ForWriting, True, TristateFalse)
 f.writeline "<html>"
 f.writeline "<h1><div align=""center"">Configura&ccedil;&atilde;o dos usu&aacuterios - Paulo Sant´anna - contato@paulosantanna.com</div></h1>"
 f.writeline( "<h1>Nome do Dominio: " & prov.name & "</h1>")
 f.writeline "<table cellspacing=""2"" cellpadding=""2"" border=""1"" align=""center"" valign=""middle"">"
 f.writeline "<tr>"
 f.writeline "    <td>Nome do usu&aacuterio</td>"
 f.writeline "    <td>Nome completo</td>"
 f.writeline "    <td>descri&ccedil;&atilde;o</td>"
 f.writeline "    <td>Conta bloqueada ?</td>"
 f.writeline "    <td>Letra da Pasta base</td>"
 f.writeline "    <td>Pasta base</td>"
 f.writeline "    <td>Script de Login</td>"
 f.writeline "    <td>Membro de:</td>"
 f.writeline "</tr>"
 
For each o in prov
   
  If o.Class = "User" Then


    if o.name = "" then
        f.writeline ("<tr>   <td>  </td>")
    else
	f.writeline ("<tr> <td>" & o.name & "</td>")
    end if
    if o.FullName = "" then
        f.writeline ("<td>   </td>")
    else
	f.writeline ("<td>" & o.FullName & "</td>")
    end if
    if o.Description = "" then
        f.writeline ("<td>   </td>")
    else
	f.writeline ("<td>" & o.Description & "</td>")
    end if
    if o.IsAccountLocked = true then
       f.writeline ("<td> Sim </td>")
    else
       f.writeline ("<td> N&atilde;o </td>")
    end if
    if o.HomeDirDrive = "" then
        f.writeline ("<td>   </td>")
    else
	f.writeline ("<td>" & o.HomeDirDrive & "</td>")
    end if
    if o.HomeDirectory = "" then
        f.writeline ("<td>   </td>")
    else
	f.writeline ("<td>" & o.HomeDirectory & "</td>")
    end if    
    if o.LoginScript = "" then
        f.writeline ("<td>   </td>")
    else
	f.writeline ("<td>" & o.LoginScript & "</td>")
    end if    
    members = ""     	

	For Each group In o.groups
        
		if members<> "" then
			members= members+ ",  "
		end if
		members= members + "  " + group.Name 

     	Next
	if members<> "" then
		members= members+ "."
		f.writeline("<td>" & members & "</td>")
	else
		f.writeline("<td> </td>")
	end if


  End If

 Next
 
f.writeline "</table>"
f.writeline "</html>"

f.close
