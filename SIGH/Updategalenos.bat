rem @echo off
copy "\\172.16.3.202\instala\GalenhosUpdate\"  "C:\Program Files (x86)\Digital Works Corporation\GalenHos"
copy "\\172.16.3.202\instala\GalenhosUpdate\Archivos"  "C:\Program Files (x86)\Digital Works Corporation\GalenHos\Archivos"
copy "\\172.16.3.202\instala\GalenhosUpdate\Imagenes"  "C:\Program Files (x86)\Digital Works Corporation\GalenHos\Imagenes"
copy "\\172.16.3.202\instala\GalenhosUpdate\Plantillas"  "C:\Program Files (x86)\Digital Works Corporation\GalenHos\Plantillas"
copy "\\172.16.3.202\instala\GalenhosUpdate\"  "C:\Windows\System32"
copy "\\172.16.3.202\instala\GalenhosUpdate\"  "C:\Windows\SysWOW64"

regsvr32.exe /s sighcomun.dll
regsvr32.exe /s sighdatos.dll
regsvr32.exe /s sighFacturacion.dll
regsvr32.exe /s sighnegocios.dll
regsvr32.exe /s sighproxies.dll
regsvr32.exe /s sighreportes.dll
regsvr32.exe /s sighFarmacia.dll
regsvr32.exe /s sighImagen.dll
regsvr32.exe /s sighLaboratorio.dll
regsvr32.exe /s duzactx.dll
regsvr32.exe /s dzactx.dll
regsvr32.exe /s mschrt20.ocx
regsvr32.exe /s igscroll40.ocx
regsvr32.exe /s ssInput1.ocx
regsvr32.exe /s OWC11.DLL
regsvr32.exe /s msxml.dll
regsvr32.exe /s sighEntidades.dll
regsvr32.exe /s sighCatalogos.dll
regsvr32.exe /s DllReg.dll
regsvr32.exe /s SIGHhisDigitacion.dll
regsvr32.exe /s SIGHsis.dll
regsvr32.exe /s crviewer.dll
regsvr32.exe /s SIGHIntegracion.dll
regsvr32.exe /s pvxplore8.ocx
regsvr32.exe /s scrrun.dll