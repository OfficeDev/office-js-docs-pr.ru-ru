# <a name="requestedheight-element"></a>Элемент RequestedHeight

Указывает исходную высоту окна контентной или почтовой надстройки (в пикселях). 

**Тип надстройки:** контентные и почтовые надстройки

## <a name="syntax"></a>Синтаксис

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>Родительские элементы

- [DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000
- [DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450
- [ExtensionPoint](extensionpoint.md)  (контекстные почтовые надстройки) со значением в диапазоне от 140 до 450 для точки расширения **DetectedEntity** и в диапазоне от 32 до 450 для точки расширения **CustomPane**