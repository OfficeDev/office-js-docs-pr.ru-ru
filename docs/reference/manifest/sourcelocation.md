# <a name="sourcelocation-element"></a>Элемент SourceLocation

Указывает расположения исходного файла для надстройки Office как URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.

**Тип надстройки:** содержимое, область задач, почта

## <a name="syntax"></a>Синтаксис

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>Содержится в

- [DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)
- [FormSettings](formsettings.md) (почтовые надстройки)
- [ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки)

## <a name="can-contain"></a>Может содержать

[Переопределение](override.md)

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|обязательный|Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).|
