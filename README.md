# orabooster-epplus

A fork of **EPPlusFree 4.5.3.8** used by [OraBooster](https://www.orascope.net), with a faster save path.
It is not an official EPPlus release and is not supported by EPPlus Software AB or by the EPPlusFree author.

## Origin

| | Repository | Version | License |
|---|---|---|---|
| Original | https://github.com/JanKallman/EPPlus (archived) | EPPlus 4.5.3.3, the last LGPL version | LGPL |
| Continued by | https://github.com/rimland/EPPlus | EPPlusFree 4.5.3.8 (commit c777bbe) | LGPL-3.0 |
| This fork | https://github.com/orascope/orabooster-epplus | 4.5.3.8-ora.N | LGPL-3.0 |

EPPlus 5 and later are under a different (Polyform Noncommercial) license and are not used here.

## License

GNU Lesser General Public License v3.0: [COPYING.LESSER](COPYING.LESSER), with the GNU GPL v3.0 it refers to in [COPYING](COPYING).
The copyright notices of the original authors are kept in the source files.

## Changes in this fork

See [CHANGES.md](CHANGES.md). Every change is also marked in the code with `orabooster-epplus:`.

## Build

```
dotnet pack EPPlus/EPPlus.MultiTarget.csproj -c Release
```

Produces `OraBooster.EPPlus.<version>.nupkg` (assembly `OraBooster.EPPlus.dll`, namespace `OfficeOpenXml` as before).
`tools/SaveBench` measures saving and writes sample workbooks; `tools/SaveBench/compare-xlsx.ps1` compares two workbooks part by part.

---

The README of EPPlusFree follows.
# Continuous Version of EPPlus Free Edition 4.5.3.3
**The [JanKallman/EPPlus](https://github.com/JanKallman/EPPlus) repository has been archived by the owner. So I fork it to my repository and continue to maintain the source code.**

---
**This repository has moved to https://github.com/EPPlusSoftware/EPPlus.** 

**The code in this archive represents the final version of EPPlus under LGPL. There will be no more activity here.**  

EPPlus will from version 5 switch license from **LGPL** to [Polyform Noncommercial 1.0.0]( https://polyformproject.org/licenses/noncommercial/1.0.0/) license.  
With the new license EPPlus is still free to use in some cases, but will require a commercial license to be used in a commercial business.

More information on the license change on [our website]( https://www.epplussoftware.com)
***
Create advanced Excel spreadsheets using .NET, without the need of interop.

EPPlus is a .NET library that reads and writes Excel files using the Office Open XML format (xlsx). 
EPPlus has no dependencies other than .NET.
 
## EPPlus supports:
* Cell Ranges 
* Cell styling (Border, Color, Fill, Font, Number, Alignments) 
* Data validation 
* Conditional formatting 
* Charts 
* Pictures 
* Shapes 
* Comments 
* Tables 
* Pivot tables 
* Protection 
* Encryption 
* VBA 
* Formula calculation 
* Many more... 

## Overview
This project started with the source from ExcelPackage. It was a great project to start from.
It had the basic functionality needed to read and write a spreadsheet.
Advantages over other:
EPPlus uses dictionaries to access cell data, making performance a lot better.
Complete integration with .NET 

## Support
All support is currently referred to [Stack overflow](https://stackoverflow.com/questions/tagged/epplus). 
A tutorial is available in the wiki and the sample project can be downloaded with each version. 
The old site at [Codeplex](http://epplus.codeplex.com) also contains material that can be helpful. 
Bugs and new feature requests can be added to the issues tracker. 

## License
The project is licensed under the GNU Library General Public License (LGPL). 
