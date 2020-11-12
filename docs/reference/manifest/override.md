---
title: Элемент Override в файле манифеста
description: Элемент override позволяет указать значение параметра в зависимости от указанного условия.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996314"
---
# <a name="override-element"></a>Элемент Override

Предоставляет способ переопределения значения параметра манифеста в зависимости от указанного условия. Существует два типа условий:

- Языковой стандарт Office, отличный от используемого по умолчанию.
- Шаблон поддержки набора требований, отличный от шаблона по умолчанию.

Существует два типа элементов: `<Override>` один — для переопределения языкового стандарта, который называется **локалетокеноверриде** , а другой — для переопределения набора требований, именуемого **рекуиременттокеноверриде**. Но `type` для элемента нет параметров `<Override>` . Разница определяется родительским элементом и типом родительского элемента. `<Override>`Элемент, который находится внутри `<Token>` элемента `xsi:type` , который `RequirementToken` должен иметь тип **рекуиременттокеноверриде**. `<Override>`Элемент внутри любого другого родительского элемента или внутри `<Override>` элемента типа `LocaleToken` должен иметь тип **локалетокеноверриде**. Каждый тип описывается в отдельных разделах ниже.

## <a name="override-element-of-type-localetokenoverride"></a>Элемент override элемента типа Локалетокеноверриде

`<Override>`Элемент выражает условное значение и может быть прочитано как "If... Then... " Оператор. Если `<Override>` элемент имеет тип **локалетокеноверриде** , `Locale` атрибут является условием, а `Value` атрибут — консекуент. Например, прочтите следующий текст: "при настройке языкового стандарта Office fr-FR отображается имя" Лектеур видéо "."

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

### <a name="syntax"></a>Синтаксис

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>Содержится в

|Элемент|
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
|[Маркер](token.md)|

### <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Языковой стандарт|string|Обязательный|Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.|
|Значение|string|Обязательный|Задает значение параметра, представленное для указанного языкового стандарта.|

### <a name="examples"></a>Примеры

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a>См. также

- [Локализация надстроек для Office](../../develop/localization.md)
- [Сочетания клавиш](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a>Элемент override элемента типа Рекуиременттокеноверриде

`<Override>`Элемент выражает условное значение и может быть прочитано как "If... Then... " Оператор. Если `<Override>` элемент имеет тип **рекуиременттокеноверриде** , дочерний `<Requirements>` элемент выражает условие, а `Value` атрибут — консекуент. Например, первое `<Override>` в приведенном ниже примере считывается, если текущая платформа поддерживает феатуреоне версии 1,7, а затем используйте строку "олдаддинверсион" вместо `${token.requirements}` маркера в URL-адресе в URL-адресе "бабушке" `<ExtendedOverrides>` (вместо строки по умолчанию "Upgrade"). "

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

**Тип надстройки:** надстройки области задач

### <a name="syntax"></a>Синтаксис

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>Содержится в

|Элемент|
|:-----|
|[Маркер](token.md)|

### <a name="must-contain"></a>Должен содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Requirements](requirements.md)|||x|

### <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Значение|string|Обязательный|Значение маркера "бабушке" при удовлетворении условия.|

### <a name="example"></a>Пример

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание элемента Requirements в манифесте](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [Сочетания клавиш](../../design/keyboard-shortcuts.md)
