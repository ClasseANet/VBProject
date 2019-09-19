Set wshShell = CreateObject("WScript.Shell")
Set PPath=wshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

COPY  "C:\Program Files (x86)\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Program Files (x86)\ClasseA\Admin\DLL\G3RREV2.ZIA"
del "C:\Program Files (x86)\ClasseA\Admin\DLL\G3RREV.ZIA"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Sistemas\Dsr\Projeto3R\Script\Revisao\*"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Sistemas\Dsr\Projeto3R\Script\3R06_InsertMenu.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Sistemas\Dsr\Projeto3R\Script\3R07_Pesquisas.Sql"

del "C:\Sistemas\Dsr\Projeto3R\Setup\DEP.zip"
Copy C:\WINDOWS\system32\ClasseA\VersaoFTP.dll "C:\Sistemas\Dsr\Projeto3R\Setup\DEP\VersaoFTP.dll"
Copy C:\WINDOWS\system32\ClasseA\xLib.dll "C:\Sistemas\Dsr\Projeto3R\Setup\DEP\xLib.dll"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Sistemas\Dsr\Projeto3R\Setup\DEP.zip" "C:\Sistemas\Dsr\Projeto3R\Setup\DEP\*"

del "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\Result01.txt"
del "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R01_Banco.sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R02_Users.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R03_Defaults.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R04_Tabelas.sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R05_Insert.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R06_InsertMenu.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R07_Pesquisas.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\GO_3R03_Defaults.Sql"
"C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex "C:\Program Files (x86)\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" "C:\Sistemas\Dsr\Projeto3R\Script\GO_3R05_Delete.Sql"

REM del %programfiles%\"ClasseA\Admin\Instalacao\Setup\P3R.zia"
REM Move C:\Sistemas\VBInstaller\Output\P3R\DISK_1\P3R.msi %programfiles%\"ClasseA\Admin\Instalacao\P3R.msi"
REM "C:\Program Files\WinZip\winzip64.exe"  -a -r -ef -ex %programfiles%\"ClasseA\Admin\Instalacao\Setup\P3R.zia" %programfiles%\"ClasseA\Admin\Instalacao\Setup\"
REM "C:\Program Files\WinZip\winzip64.exe"  -a -ef -ex %programfiles%\"ClasseA\Admin\Instalacao\Setup\P3R.zia" %programfiles%\"ClasseA\Admin\Instalacao\P3R.msi"



