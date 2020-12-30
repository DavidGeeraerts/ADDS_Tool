# Active Directory Functional Level

[Forest and Domain Functional Levels documentation](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-adts/564dc969-6db3-49b3-891a-f2f8d0a68a7f?redirectedfrom=MSDN)

| Domain Functional Levels (domainFunctionality)	| Forest Functional Levels(forestFunctionality)	|
|:--------------------------------------------------|:-------------------------------------------------|
| 0 - Windows Server 2000 mixed						| 0 — Windows 2000
| 0 - Windows Server 2000 native					| 0 — Windows 2000
| 1 - Windows Server 2003 interim					| 1 — Windows Server 2003 interim
| 2 - Windows Server 2003							| 2 — Windows Server 2003
| 3 - Windows Server 2008							| 3 - Windows Server 2008
| 4 - Windows Server 2008 R2 domain level			| 4 - Windows Server 2008 R2 domain level
| 5 - Windows Server 2012							| 5 - Windows Server 2012
| 6 - Windows Server 2012 R2						| 6 - Windows Server 2012 R2
| 7 - Windows Server 2016							| 7 - Windows Server 2016


`DSQUERY * "DC=domain,DC=root" -scope base -attr msDS-Behavior-Version`