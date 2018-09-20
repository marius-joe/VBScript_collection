# VBScript_collection
[![VBScript version](https://img.shields.io/badge/VBScript-5.8-blue.svg)](https://www.w3schools.com/asp/asp_ref_vbscript_functions.asp)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-brightgreen.svg)](https://github.com/marius-joe/VBScript_collection/graphs/commit-activity)
[![GitHub issues](https://img.shields.io/github/issues/marius-joe/VBScript_collection.svg)](https://github.com/marius-joe/VBScript_collection/issues/)
[![GitHub license](https://img.shields.io/github/license/marius-joe/VBScript_collection.svg)](https://github.com/marius-joe/VBScript_collection/blob/master/LICENSE)

*[coming soon]*

**Collection of helper/utility scripts**

For Windows system administrators, Microsoft suggests migrating to PowerShell. However, the VBScript engine will continue to be shipped with future releases of Microsoft Windows.
<br/>
<br/>
**My advice:** Don't create whole new complex scripts using VBScript anymore. Using and maintaining old stable ones is fine.
<br/>
<br/>
**Why use it for small scripts ?:**<br/>
(normally!) you cannot execute a Powershell script by clicking and have to invoke the script by<br/>
*PowerShell.exe -NoProfile -ExecutionPolicy Bypass -file mypowerscript.ps1*<br/>
That's "Security by design", the Powershell team advise against changing ps1 files to run on double click)<br/>
-> not comfortable if you just want to execute some small scripts<br/>
-> VBScripts can be run by clicking on any machine with default configuration
