# axWidgetC
VB6 Widget UserControl, version modificada del control Firenze Label de Martin Vartiak publicada en [VBForum]

## Prerequisitos

Este control tiene las siguientes dependecias:
```
* vbRichClient5.dll
* vb_cairo_sqlite.dll
```
las que encontrarás en este mismo repositorio, estos archivos deben ir en la misma carpeta que el OCX o en la Carpeta raiz de tu proyecto si prefieres usar el UserControl sin compilar.

### Instalación

Para usar el control puedes descargar la ultima release (OCX) liberada en este repositorio y copiarla a la carpeta Windows/System si tu sistema es x32 o Windows/SysWOW64 si tu sistema es x64, y registrarlo con REGSVR32

Windows 32bit
```
C:\Windows\System\Regsvr32 axWidgetc.ocx
```

Windows 64bit
```
C:\Windows\SysWOW64\Regsvr32 axWidgetc.ocx
```

No debes olvidar copiar los archivos: vbRichClient5.dll y vb_cairo_sqlite.dll, igualmente en la misma carpeta que el OCX y registralos de la misma forma

Windows 32bit
```
C:\Windows\System\Regsvr32 vbRichClient5.dll
C:\Windows\System\Regsvr32 vb_cairo_sqlite.dll
```

Windows 64bit
```
C:\Windows\SysWOW64\Regsvr32 vbRichClient5.dll
C:\Windows\SysWOW64\Regsvr32 vb_cairo_sqlite.dll
```

Ya con eso el control estará disponible para su uso con VB6.


