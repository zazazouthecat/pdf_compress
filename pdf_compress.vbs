'Objects
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set wShell=CreateObject("WScript.Shell")
'scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
scriptdir=wShell.CurrentDirectory

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' check si le dossier compresssed existe, si il n existe pas on le creer
path="compressed"
exists = oFSO.FolderExists(path)
if (exists) then 
	'nothing
	else
		oFSO.CreateFolder "compressed"
end if
''''''''''''''''''''''''''''''''''''''''''''''''''

'' Boite de dialogue pour choisir le fichier ''
Set oExec=wShell.Exec("mshta.exe ""about:</head><meta http-equiv='X-UA-Compatible' content='IE=EmulateIE11'/></head><input type=file id=FILE name=files><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine

'' Si aucun fichier choisi on quitte le script
if IsEmpty(sFileSelected) OR IsNull(sFileSelected) or sFileSelected="" Then
	wscript.Quit
end if
''''''''''''''''''''''''''''''''''''''''''''''''

'' On rÃ©cupÃ¨re le nom du fichier uniquement ''
pos = InstrRev (sFileSelected, "\")
pos = pos + 1
outfilename=mid(sFileSelected,pos)
''''''''''''''''''''''''''''''''''''''''''''''

'' On appel Ghostscript pour compresser le pdf ''
'wscript.echo sFileSelected
'Set oExec=wShell.Exec("C:\Program Files (x86)\gs\gs9.21\bin\gswin32c.exe -dNOPAUSE -dBATCH -dSAFER -dPDFSETTINGS=/ebook -dCompatibilityLevel=1.4 -sDEVICE=pdfwrite -sOutputFile=" & """compressed\"""& sFileSelected & """" & sFileSelected &"""")
Set oExec=wShell.Exec("gswin32.exe -dNOPAUSE -dBATCH -dSAFER -dPDFSETTINGS=/ebook -dCompatibilityLevel=1.4 -sDEVICE=pdfwrite -sOutputFile=" & """compressed\" & outfilename  & """ """ & sFileSelected &"""")
''''''''''''''''''''''''''''''''''''''''''''''''''

'' on ouvre le dossier de resultat contenant les fichiers compressÃ© pour l'utilisateur ''

'' Tant que Ghostscript n'a pas terminÃ©, on attend avant d'ouvrir le dossier de resultat
Do While oExec.Status = 0
WScript.Sleep 10
Loop
'wscript.echo "EXPLORER.exe """ & scriptdir & "\compressed"""
wShell.run "EXPLORER.exe """ & scriptdir & "\compressed"""
'''''''''''''''''''''''''''''''''''''''''''''''''
