# <a name="override-element"></a>Элемент Override

Предоставляет способ указать значение параметра для дополнительного языкового стандарта.

**Тип надстройки:** содержимое, область задач, почта

## <a name="syntax"></a>Синтаксис

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Содержится в

|**Элемент**|
|:-----|
|[CitationText](citationtext.md)|
|[Описание](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|Locale|string|обязательный|Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых меток BCP 47, например `"en-US"`.|
|Значение|string|обязательный|Задает значение параметра, представленное для указанного языкового стандарта.|

## <a name="see-also"></a>См. также

- [Локализация надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
