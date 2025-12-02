@ECHO OFF
chcp 65001
set /p version=Enter the version (for example, 1.0.0.0):

C:\Users\bushk\.nuget\packages\ilrepack\2.0.44\tools\ILRepack.exe /out:"OpenXMLight.dll" "F:\тестовые проекты\ConsoleApp1\OpenXMLight\bin\Debug\net6.0\OpenXMLight.dll" C:\Users\bushk\.nuget\packages\documentformat.openxml\3.3.0\lib\netstandard2.0\DocumentFormat.OpenXml.dll C:\Users\bushk\.nuget\packages\documentformat.openxml.framework\3.3.0\lib\net6.0\DocumentFormat.OpenXml.Framework.dll C:\Users\bushk\.nuget\packages\system.io.packaging\8.0.1\lib\net6.0\System.IO.Packaging.dll /ver:%version%