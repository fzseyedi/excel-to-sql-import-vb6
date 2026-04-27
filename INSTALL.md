## Runtime dependencies

This application is built with Visual Basic 6 and may require the following runtime components on the target machine:

- Visual Basic 6 Runtime
- Microsoft ADO / MDAC
- Microsoft Scripting Runtime
- Microsoft Windows Common Controls (`MSCOMCTL.OCX`)
- Microsoft FlexGrid Control (`MSFLXGRD.OCX`)
- Microsoft Tabbed Dialog Control (`TABCTL32.OCX`)
- Microsoft Common Dialog Control (`COMDLG32.OCX`)

Optional components:

- Microsoft Excel for `.xlsx` export using Excel Automation
- Microsoft Jet/ACE OLEDB provider for Excel import/export through ADO/OLEDB

CSV export does not require Microsoft Excel.

This repository does not include Microsoft runtime DLL/OCX files.  
Install required Microsoft components from official sources or use a proper installer package.