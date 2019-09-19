del "C:\Arquivos de programas\ClasseA\Admin\DLL\G3RREV.ZIA"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Sistemas\Dsr\Projeto3R\Script\Revisao\*"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Sistemas\Dsr\Projeto3R\Script\3R06_InsertMenu.Sql"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\DLL\G3RREV.ZIA" "C:\Sistemas\Dsr\Projeto3R\Script\3R07_Pesquisas.Sql"

del C:\Sistemas\Dsr\Projeto3R\Setup\DEP.zip
Copy C:\WINDOWS\system32\ClasseA\VersaoFTP.dll C:\Sistemas\Dsr\Projeto3R\Setup\DEP\VersaoFTP.dll
Copy C:\WINDOWS\system32\ClasseA\xLib.dll C:\Sistemas\Dsr\Projeto3R\Setup\DEP\xLib.dll
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Sistemas\Dsr\Projeto3R\Setup\DEP.zip" "C:\Sistemas\Dsr\Projeto3R\Setup\DEP\*"

del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\Result01.txt"
del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia"
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R01_Banco.sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R02_Users.Sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R03_Defaults.Sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R04_Tabelas.sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R05_Insert.Sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R06_InsertMenu.Sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\3R07_Pesquisas.Sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\GO_3R03_Defaults.Sql
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.zia" C:\Sistemas\Dsr\Projeto3R\Script\GO_3R05_Delete.Sql


del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P01.zia"
del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P02.zia"
del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P03.zia"
del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P04.zia"
del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P05.zia"
del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P99.zia"
Move C:\Sistemas\VBInstaller\Output\P3R\DISK_1\P3R.msi "C:\Arquivos de programas\ClasseA\Admin\Instalacao\P3R.msi"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\MSI.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\P3R.msi"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P01.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\01 Windows Installer\"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P02.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\02 FrameWork\"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P03.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\03 Sql2005Express\"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P04.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\04 SqlManager\"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P05.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\05 TeamViewer\"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P99.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.mdf"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P99.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\99 Database\G3R.ldf"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P99.zia" "C:\Sistemas\Dsr\Projeto3R\Script\3R02_Users.Sql"


del "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P3R.zia"
rem Move C:\Sistemas\VBInstaller\Output\P3R\DISK_1\P3R.msi "C:\Arquivos de programas\ClasseA\Admin\Instalacao\P3R.msi"
"C:\Arquivos de programas\WinZip\winzip32.exe"  -a -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P3R.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\P3R.msi"
rem "C:\Arquivos de programas\WinZip\winzip32.exe"  -a -r -ef -ex "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\P3R.zia" "C:\Arquivos de programas\ClasseA\Admin\Instalacao\Setup\"



