'----------------------------------------------------------
' Plugin for OCS Inventory NG 2.x
' Web : http://www.ocsinventory-ng.org
' Script : Retrieve SQL Server databases of the station
' Version : 1.50
' Date : 23/03/2018
' Author : Sylvie COZIC
' Contributors : Frank BOURDEAU (OCS Inventory NG) and Stephane PAUTREL (acb78.com)
'----------------------------------------------------------
' OS checked [X] on	32b	64b	(Professionnal edition)
'	Windows XP		[ ]
'	Windows Vista	[X]	[X]
'	Windows 7		[X]	[X]
'	Windows 8.1		[X]	[X]
'	Windows 10		[X]	[X]
'	Windows 2k8R2		[X]
'	Windows 2k12R2		[X]
'	Windows 2k16		[X]
' ---------------------------------------------------------
' NOTE : No checked on Windows 8
' ---------------------------------------------------------
On Error Resume Next

'-------------------------------------------------------------------------------
' Liste les bases de données SQL Server du poste
'  4 données sont remontées :
'  - strSQLName :      Nom long du produit SQL Server
'                            Par exemple : "Microsoft SQL Server 2008 R2"
'  - strServiceName : Nom de l'instance
'                            Par exemple : "MSSQLSERVER"
'  - strEdition :      Edition.
'                            Par exemple : "Enterprise Edition (64-bit)"
'  - strVersion :      Version "chiffrée".
'                            Par exemple : "8.00.194"

Const PluginAuthor  = "Sylvie Cozic"
Const PluginDate    = "23/03/2018"
Const PluginVersion = "1.5.0"

' Historique :
' 1.0.0 - 21/10/2011 - Sylvie Grimonpont : 
'         Première version
'
' 1.1.0 - 17/10/2012 - Sylvie Grimonpont :
'         Ajout des versions Service pack de SQL Server 2008
'
' 1.2.0 - 27/09/2013 - Sylvie Grimonpont : 
'         Ajout des versions Service pack de SQL Server 2008 R2 + SQL Server 2012
'          + recalcule d'un numéro de version "officiel" SQL Server (celui qu'on retrouve en lancant "Select @@version")
'          + ecriture dans la table softwares (fonctionne à partir de l'agent OCS 2.1.0)
'
' 1.3.0 - 26/11/2013 - Sylvie Grimonpont :
'         Si aucun service SQL n'est detecté mais que des produits SQL sont installés, on remonte l'information
'
' 1.4.0 - 04/06/2015 - Sylvie Cozic (Grimonpont) :
'         Ajout de la version SQL Server 2014
'          + Ré-écrtiture complète de la partie Errorlog (quand les classes WMI ne donnent rien)
'          + Ajout du plugin dbinstances dans les logiciels
'
' 1.5.0 - 23/03/2018 - Frank Bourdeau
'         Ajout de la version SQL Server 2016 et SQL Server 2017
'
' This code is open source and may be copied and modified as long as the source
' code is always made freely available.
' Please refer to the General Public Licence http://www.gnu.org/ or Licence.txt
'-------------------------------------------------------------------------------

' DECLARATIONS

'Déclaration des constantes
Const DblQuote  = """"
Const ForReading = 1
Const HKEY_LOCAL_MACHINE = &H80000002
Const strMSSQLServerRegKey = "SOFTWARE\Microsoft\MSSQLServer\MSSQLServer\Parameters"
Const REG_SZ = 1
Const adVarChar = 200
Const MaxCharacters = 255

' RegExp pour reconstituer le n° de version "officiel" Microsoft SQL Server
' Par exemple Microsoft SQL Server 2008 10.3.5500.0 = Microsoft SQL Server 2008 10.00.5500.0 SP3 (le .3. = SP3...)
Set regexpOfficialVersion = New RegExp
regexpOfficialVersion.IgnoreCase = True
regexpOfficialVersion.Global = True
regexpOfficialVersion.Pattern = "(\d*)\.(\d*)\.(.*)"

' RegExp pour détection produits SQL Server
Set regexpSQLProduct = New RegExp
regexpSQLProduct.IgnoreCase = True
regexpSQLProduct.Global = True
regexpSQLProduct.Pattern = "^Microsoft([^ -~]| )+SQL([^ -~]| )+Server([^ -~]| )+(\d|\.)+( (R|V)\d)*( \(64-bit\))*$"

' Clefs de registre des logiciels installés à parcourir
Set objUninstallPaths = CreateObject("Scripting.Dictionary")
objUninstallPaths.Add "1", "Software\Microsoft\Windows\CurrentVersion\Uninstall"
objUninstallPaths.Add "2", "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

' Dictionnaire des logiciels installés (pour gestion des doublons de nom logiciel)
Set dicInstalledSoftwares = CreateObject("Scripting.Dictionary")
dicInstalledSoftwares.RemoveAll
dicInstalledSoftwares.CompareMode = 0


' MAIN

' Remontée des informations du plugin dans la table SOFTWARES
arrScriptName = Split(Wscript.ScriptName,".")
Result = "<SOFTWARES>" & VbCrLf
Result = Result & "<PUBLISHER>Sylvie Cozic</PUBLISHER>" & VbCrLf
Result = Result & "<NAME>" & arrScriptName(0) & "</NAME>" & VbCrLf
Result = Result & "<VERSION>" & PluginVersion & "</VERSION>" & VbCrLf
Result = Result & "<COMMENTS>Data out of OCS Plugin</COMMENTS>" & VbCrLf
Result = Result & "</SOFTWARES>"
WScript.Echo Result

' Recherche d'un service ayant sqlservr.exe dans son path. Si ce service n'existe pas, aucune base sql ne tourne.
Set objWMIcimv2 = GetObject("winmgmts:root\cimv2")
If Err = 0 Then

    Set colServices = objWMIcimv2.ExecQuery("Select Name , PathName from Win32_Service Where PathName Like '%sqlservr.exe%'")
    If Err = 0 Then
        If colServices.count > 0 Then
            If Err = 0 Then
                'WScript.Echo "SQL Server trouvé !"
                For Each objService in colServices
                    ' ServiceName
                    strServiceName = objService.Name
                    If InStr(strServiceName,"$") > 0 Then
                        strServiceName =  Mid(strServiceName, InStr(strServiceName,"$") + 1)
                    End If

                    ' PathName, Drive et Path
                    strPathName = objService.PathName
                    If Left(strPathName, 1) = DblQuote Then strPathName= Mid(strPathName, 2)
                    strDrive = Mid(strPathName, 1, 2)
                    strPath  = Mid(strPathName, 3, InStr(1, strPathName, "sqlservr.exe") - 3)

                    ' Recherche la version du fichier sqlservr.exe
                    strCIMDatafile = "Select FileName,Extension,Version from CIM_Datafile" _
                                        & " Where Drive = '" & strDrive & "'" _
                                        & " and Path = '" & Replace( strPath ,"\","\\") & "'" _
                                        & " and FileName = 'sqlservr' and Extension = 'exe'"

                    ' Si on a eu une erreur entre temps, on efface
                    Err.Clear

                    Set colSQLFile = objWMIcimv2.ExecQuery ( strCIMDatafile )
                    If Err = 0 Then
                        For Each objSQLFile in colSQLFile
                            ' WScript.Echo "Version objSQLFile (l. 131) : " & objSQLFile.Version
                            If Not IsNull(objSQLFile.Version) Then
                                arrSQLFileVersion=Split(objSQLFile.Version,".")
                                strSQLFileVersion=Cint(arrSQLFileVersion(1))
                            Else
                                strSQLFileVersion=0
                            End If

                            ' Initialisation
                            strVersion =""
                            strEdition =""
                            strServicePack = ""
                            strWMIsql = ""
                            strError = ""

                            ' Positionne la classe WMI SqlServer en fonction de la version du fichier sqlservr.exe
                            If strSQLFileVersion = 90 Then strWMIsql = "ComputerManagement"
                            If strSQLFileVersion > 90 Then strWMIsql = "ComputerManagement10"
                            If strSQLFileVersion = 110 Then strWMIsql = "ComputerManagement11"
                            If strSQLFileVersion = 120 Then strWMIsql = "ComputerManagement12"
                            If strSQLFileVersion = 130 Then strWMIsql = "ComputerManagement13"
                            If strSQLFileVersion = 140 Then strWMIsql = "ComputerManagement14"

                            ' Recherche la version et l'édition de la base SQL via la classe WMI SqlServer si disponible
                            If (strWMIsql <> "") Then
                                ' Si on a eu une erreur entre temps, on efface
                                Err.Clear
                                Set objWMIsql = GetObject("winmgmts:root\Microsoft\SqlServer\" & strWMIsql)
                                If Err = 0 Then
                                    For Each prop in objWMIsql.ExecQuery("select * from SqlServiceAdvancedProperty Where SQLServiceType = 1 And ServiceName='" & objService.Name & "'")
                                        If Err = 0 Then
                                            If (prop.PropertyName="VERSION") Then strVersion = prop.PropertyStrValue
                                            If (prop.PropertyName="SKUNAME") Then strEdition = prop.PropertyStrValue
                                        Else
                                            WriteError()
                                        End If
                                    Next
                                    If Err <> 0 Then WriteError()
                                Else
                                    WriteError()
                                End If
                            End If

                            ' Si on n'a pas pu determiner la version et l'édition à partir de la classe WMI (cas SQL Server 2000), on essaie de la déterminer à partir du fichier ERRORLOG (si en local)
                            If (((strWMIsql= "") Or (strVersion = "") Or (strEdition = "")) And (strSourceServer = ".")) Then

                                ' On ne trappe plus les erreurs...
                                On Error Goto 0

                                ' Init
                                strErrorlogPath = ""
                                Set objFSO = CreateObject("Scripting.FileSystemObject")
                                
                                ' Recherche le répertoire le répertoire des ERRORLOG de la base SQL à partir du chemin de l'exécutable du service
                                If objFSO.FolderExists(strDrive & strPath) Then
                                     strServiceParentPath = objFSO.GetParentFolderName(strDrive & strPath)
                                     If objFSO.FolderExists(strServiceParentPath & "\LOG") Then
                                          strErrorlogPath = strServiceParentPath & "\LOG"
                                     End If
                                End If

                                ' Tente la recherche le répertoire des ERRORLOG dans la base de registre
                                If strErrorlogPath = "" Then
                                     Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")
                                     If objRegistry.EnumKey (HKEY_LOCAL_MACHINE, strMSSQLServerRegKey, arrSubKeys) = 0 Then
                                         strErrorlogFile = ""
                                         objRegistry.EnumValues HKEY_LOCAL_MACHINE, strMSSQLServerRegKey, arrValueNames, arrValueTypes
                                         For I=0 To UBound(arrValueNames)
                                              If arrValueTypes(I) = REG_SZ Then
                                                  objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strMSSQLServerRegKey,arrValueNames(I),strValue
                                                  ' Dans HKLMSOFTWARE\Microsoft\MSSQLServer\MSSQLServer\Parameters, le paramètre qui commence par "-e" définie le path de ERRORLOG
                                                  If Left(strValue,2) = "-e" Then
                                                      strErrorlogFile = Mid(strValue,3)
                                                      Exit For
                                                  End If
                                              End If
                                         Next

                                         ' Si on a trouvé le chemin du fichier ERRORLOG dans la base de registre, on récupère le répertoire de celui-ci
                                         If strErrorlogFile <> "" Then
                                             ' Teste l'existence du fichier
                                             If objFSO.FileExists(strErrorlogFile) Then
                                                strErrorlogPath = objFSO.GetParentFolderName(strErrorlogFile)
                                             Else
                                                 strError = "Le fichier " & strErrorlogFile & " n'existe pas"
                                             End If
                                         Else
                                             strError = "Information sur fichier ERRORLOG non trouvée en base de registre"
                                         End If
                                     Else
                                         strError = "Information sur fichier ERRORLOG non trouvée en base de registre"
                                     End If
                                End If

                                ' Si on a trouvé le répertoire des ERRORLOG on essaie de lire l'un des fichiers ERRORLOG
                                If strErrorlogPath <> "" Then
                                    Set objErrorlogFolder = objFSO.GetFolder(strErrorlogPath)
                                    Set colErrorlogFiles = objErrorlogFolder.Files
                                    For Each objErrorlogFile in colErrorlogFiles
                                        If objFSO.GetBaseName(objErrorlogFile) = "ERRORLOG" Then
                                            ' Teste la taille du fichier
                                            If objErrorlogFile.Size > 0 Then
                                                strErrorlogFile = objErrorlogFile.Path
                                                ' Lit le fichier
                                                ReadErrorLog(strErrorlogFile)
                                                'Sort de la boucle
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If strVersion = "" Then
                                        strError = "Le fichier " & strErrorlogPath & "\ERRORLOG n'est pas exploitable"
                                    End If
                                End If
                            End If

                            ' Récupération du nom SQL Server
                            ' WScript.Echo "Version de SQL : " & strVersion
                            If Left(strVersion,3) = "6.5" Then
                                strSQLName = "Microsoft SQL Server 6.5"
                            ElseIf Left(strVersion,1) = "7" Then
                                strSQLName = "Microsoft SQL Server 7.0"
                            ElseIf Left(strVersion,1) = "8" Then
                                strSQLName = "Microsoft SQL Server 2000"
                            ElseIf Left(strVersion,1) = "9" Then
                                strSQLName = "Microsoft SQL Server 2005"
                            ElseIf Left(strVersion,2) = "10" Then
                                strSQLName = "Microsoft SQL Server 2008"
                                If Left(strVersion,4) = "10.5" Then
                                    strSQLName = "Microsoft SQL Server 2008 R2"
                                End if
                            ElseIf Left(strVersion,2) = "11" Then
                                strSQLName = "Microsoft SQL Server 2012"
                            ElseIf Left(strVersion,2) = "12" Then
                                strSQLName = "Microsoft SQL Server 2014"
                            ElseIf Left(strVersion,2) = "13" Then
                                strSQLName = "Microsoft SQL Server 2016"
                            ElseIf Left(strVersion,2) = "14" Then
                                strSQLName = "Microsoft SQL Server 2017"
                            Else
                                strSQLName = "Microsoft SQL Server"
                            End if

                            ' Re-"calcule" le numéro officiel de la version SQL et du service pack le cas échéant
                            ' Par exemple Microsoft SQL Server 2008 10.3.5500.0 = Microsoft SQL Server 2008 10.00.5500.0 SP3 (le .3. = SP3...)
                            Set matchesVersion = regexpOfficialVersion.Execute(strVersion)
                            If matchesVersion.Count > 0 Then
                                ' Seulement les versions > 7
                                If CInt(matchesVersion(0).Submatches(0)) > 7 Then
                                    ' Pad à gauche le 2ème nombre de la version : xx.3.xx -> xx.03.xx
                                    strPadSubversion=Right(String(2, "0") & matchesVersion(0).Submatches(1),2)
                                    ' Récupère le numéro de Service Pack si existant xx.03.xx -> SP3, xx.51.xx -> SP1,...
                                    If Right (strPadSubversion,1) <> "0" Then
                                        strServicePack = " (SP" & Right (strPadSubversion,1) & ")"
                                    End IF
                                    ' Reforme le numéro de version "officielle" : xx.03.xx -> xx.00.xx, xx.51.xx -> xx.50.xx
                                    strVersion = matchesVersion(0).Submatches(0) & "." & Left(strPadSubversion,1) & "0." & matchesVersion(0).Submatches(2)
                                End If
                            End If

                            ' Ecrit les données de sortie en XML
                            ' Les données disponibles sont :
                            ' - strSQLName :      Nom long du produit SQL Server
                            ' - strServiceName : Nom de l'instance
                            ' - strEdition :      Edition. Par exple : Enterprise Edition (64-bit) / Express Edition
                            ' - strVersion :      Version "chiffrée". Par exemple : 8.00.194 / 10.50.1600.1
                            ' Le format remonté est spécifique à un process interne. A vous d'adapter en fonction de vos besoins. :-)
                            ' On garde <DBINSTANCES> pour le côté ergonomique sur l'interface Web OCS
                            Result = "<DBINSTANCES>" & VbCrLf
                            Result = Result & "<PUBLISHER>Microsoft Corporation</PUBLISHER>" & VbCrLf
                            Result = Result & "<VERSION_NAME>" & strSQLName & strServicePack & "</VERSION_NAME>" & VbCrLf
                            Result = Result & "<VERSION>" & strVersion & "</VERSION>" & VbCrLf
                            Result = Result & "<EDITION>" & strEdition & "</EDITION>" & VbCrLf
                            Result = Result & "<INSTANCE>" & strServiceName & "</INSTANCE>" & VbCrLf
                            Result = Result & "</DBINSTANCES>"
                            WScript.Echo Result

                            ' Remontée dans la table SOFTWARES (fonctionne à partir de la version 2.1.0 de l'agent Windows)
                            Result = "<SOFTWARES>" & VbCrLf
                            Result = Result & "<PUBLISHER>Microsoft Corporation</PUBLISHER>" & VbCrLf
                            Result = Result & "<NAME>" & strSQLName & strServicePack & strEdition & "</NAME>" & VbCrLf
                            Result = Result & "<VERSION>" & strVersion & "</VERSION>" & VbCrLf
                            Result = Result & "<COMMENTS>Data out of OCS Plugin</COMMENTS>" & VbCrLf
                            Result = Result & "</SOFTWARES>"
                            WScript.Echo Result

                        Next
                    Else
                        WriteError()
                    End If
                Next
            Else
                WriteError()
            End If
        Else
            ' Aucun service SQL Server trouvé.
            ' Si des produits SQL Server sont installés, on précise dans DBInstance qu'aucun service ne tourne pour ces produits.

            'WScript.Echo "Aucun SQL Server trouvé !"
            On Error Goto 0

            'Parcours les logiciels installés
            Set objReg=GetObject("winmgmts:root\default:StdRegProv")
            colKeyPaths = objUninstallPaths.Items
            For Each strKeyPath in colKeyPaths
                objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
                If IsArray(arrSubKeys) Then
                    For Each Subkey in arrSubKeys
                        strExpKeyPath = strKeyPath & "\" & Subkey
                        objReg.EnumValues HKEY_LOCAL_MACHINE, strExpKeyPath, arrEntryNames ,arrValueTypes
                        strDisplayName = ""
                        If IsArray(arrEntryNames) Then
                            For i = 0 To UBound(arrEntryNames)
                                If InStr(1, arrEntryNames(i), "DisplayName", vbTextCompare) Then
                                    If InStr(1, arrEntryNames(i), "ParentDisplayName", vbTextCompare) Then
                                        '## Ne rien Faire
                                    Else
                                        If arrValueTypes(i) = REG_SZ Then
                                             objReg.GetStringValue HKEY_LOCAL_MACHINE, strExpKeyPath, arrEntryNames(i), strValue
                                             If strValue <> "" AND NOT IsNull(strValue) Then
                                                  strDisplayName = strValue
                                                  strDisplayName = Replace(strDisplayName, "[", "(")
                                                  strDisplayName = Replace(strDisplayName, "]", ")")
                                                  strDisplayName = Replace(strDisplayName, Chr(160), " ")
                                             End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        If strDisplayName <> "" Then
                            ' Si ce logiciel n'est pas déjà vu dans la boucle
                            If Not dicInstalledSoftwares.Exists(strDisplayName) Then
                                ' Ajout au dictionnaire des logiciels installés
                                dicInstalledSoftwares.Add strDisplayName, strDisplayName
                                ' Si ce logiciel est un produit SQL Server
                                Set matchesSQLProduct = regexpSQLProduct.Execute(strDisplayName)
                                If matchesSQLProduct.Count > 0 Then
                                    ' Ecrit les données de sortie en XML dans la table Dbinstances
                                    Result = "<DBINSTANCES>" & VbCrLf &_
                                    Result = Result & "<PUBLISHER>Microsoft Corporation</PUBLISHER>" & VbCrLf &_
                                    Result = Result & "<VERSION_NAME>" & strDisplayName & "</VERSION_NAME>" & VbCrLf &_
                                    Result = Result & "<INSTANCE>Aucun service</INSTANCE>" & VbCrLf &_
                                    Result = Result & "</DBINSTANCES>"
                                    WScript.Echo Result
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            Set dicInstalledSoftwares = Nothing
        End if
    Else
        WriteError()
    End If
Else
    WriteError()
End If

On Error Goto 0

WScript.Quit

Sub WriteError()
    strError = "Erreur " & Err.Number & " - " & Err.Description

    '            ' On écrit l'erreur dans le fichier

    Err.Clear
End Sub


Function TestFileEncoding(strFilePath)
' Retourne
'  0 si le fichier testé est en AscII
' -1 si le fichier testé est en Unicode

    ' Init
    TestFileEncoding = 0
    
    ' Ouvre le fichier et récupère les 3 premiers caractères
    Set testFile = objFSO.OpenTextFile(strFilePath)
    char1 = testFile.read(1)
    char2 = testFile.read(1)
    char3 = testFile.read(1)
    testFile.Close
    
    ' Teste les 3 premiers caractères pour voir si c'est de l'Unicode
    If (Asc(char1) = 255 And Asc(char2) = 254) Then
      TestFileEncoding = -1
    ElseIf (Asc(char1) = 254 And Asc(char2) = 255) Then
        TestFileEncoding = -1
    ElseIf (Asc(char1) = 239 And Asc(char2) = 187 And Asc(char3) = 191) Then
        TestFileEncoding = -1
    Else
      TestFileEncoding = 0
    End If

End Function

Function MultilineTrim (Byval TextData)
     Dim textRegExp
     Set textRegExp = new regexp
     textRegExp.Pattern = "\s{0,}(\S{1}[\s,\S]*\S{1})\s{0,}"
     textRegExp.Global = False
     textRegExp.IgnoreCase = True
     textRegExp.Multiline = True

     If textRegExp.Test (TextData) Then
          MultilineTrim = textRegExp.Replace (TextData, "$1")
     Else
          MultilineTrim = ""
     End If
End Function


Sub ReadErrorLog (strFilePath)
' Lit un fichier ErrorLog et récupère les informations de version et d'édition SQL

     ' RegExp pour récupération de la version dans un fichier ERRORLOG
     Set regexpVersion = New RegExp
     regexpVersion.IgnoreCase = True
     regexpVersion.Global = True
     regexpVersion.Pattern = "-[^-\(]+(\(|$)"

     ' RegExp pour récupération de l'édition dans un fichier ERRORLOG
     Set regexpEdition = New RegExp
     regexpEdition.IgnoreCase = True
     regexpEdition.Global = True
     regexpEdition.Pattern = "^.*dition|^.* on |^.*"

    ' Ouvre le fichier en fonction de son encodage
    intFileMode = TestFileEncoding(strFilePath)
    Set objTextFile =objFSO.OpenTextFile(strFilePath, ForReading, False, intFileMode)

    ' Lit le fichier
    For i = 1 To 4
        strErrorlogText = objTextFile.Readline

        ' La version se trouve sur la première ligne
        If i = 1 Then
            Set versions = regexpVersion.Execute(strErrorlogText)
            For Each version In versions
                strVersion = version
                If Left(strVersion,1)  = "-" Then strVersion = Mid(strVersion,2)
                If Right(strVersion,1) = "(" Then strVersion = Left(strVersion,Len(strVersion)-1)
                strVersion = MultilineTrim(strVersion)
            Next
            If strVersion="" Then strVersion = "Errolog"
        End If

        ' L'édition se trouve sur la quatrième ligne
        If i = 4 Then
            'strEdition = strErrorlogText
            Set editions = regexpEdition.Execute(strErrorlogText)
            For Each edition In editions
                strEdition = edition
                If Right(strEdition,4) = " on " Then strEdition = Left(strEdition,Len(strEdition)-4)
                strEdition = MultilineTrim(strEdition)
            Next
            If strEdition="" Then strEdition = "Errolog"
            
            ' Si on trouve une information de Service pack on en profite pour la remonter
            If InStr(strErrorlogText,"Service Pack") > 0 Then
                strServicePack = " (SP" & Mid(strErrorlogText, InStr(strErrorlogText,"Service Pack") + 13)
            End If
        End If

    Next

    ' Ferme le fichier
    objTextFile.Close

End Sub
