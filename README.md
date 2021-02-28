# Erebonia
A XLSX&lt;->PO converter for Sen Scripts Decompiler

Minimum requisites
-----------------------------------------------------------
.NET Core 3.0

Spanish
-----------------------------------------------------------
Abre una consola y aplica los parametros correspondientes junto a la ruta del archivo.
Uso: Erebonia.exe {modo} {archivo}
Modos: -e {archivo} {lang}: Exportar archivo .XLSX a un archivo GetText Po
       -i {archivo} {xls_original}: Importar archivo GetText Po para modificar un archivo .XLSX

English
-----------------------------------------------------------
Open a CLI and apply the arguments next to the file's path
Usage: Erebonia.exe {mode} {file}
Modes: -e {file} {lang}: Export .XLSX file to a GetText Po file
       -i {file} {original_xls}: Import a GetText Po file and outputs a .XLSX file

Credits
-----------------------------------------------------------
Credits to Pleonex for the usage of Yarhl library and EPPlusSoftware for EPPlus
(https://github.com/EPPlusSoftware/EPPlus/blob/develop/license.md)
Special thanks to Darkmet98 for his Po2BinaryEasy solution and to Twn for CS3ScenarioDecompiler
