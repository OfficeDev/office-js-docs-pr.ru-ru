---
title: Элемент Override в файле манифеста
description: Элемент Переопределения позволяет указать значение параметра в зависимости от заданного условия.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: d6f91d32a604a939118e42de882976239006c5235ea65a55d713712ebca4ee70
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093648"
---
# <a name="override-element"></a>Элемент Override

Предоставляет способ переопределения значения параметра манифеста в зависимости от указанного условия. Существует три типа условий:

- Локальный Office, который отличается от по `LocaleToken` умолчанию, называется **LocaleTokenOverride**.
- Шаблон поддержки набора требований, который отличается от шаблона по `RequirementToken` умолчанию, называемого **RequirementTokenOverride**.
- Источник отличается от по `Runtime` умолчанию, называется **RuntimeOverride**.

Элемент, который находится внутри элемента, должен `<Override>` иметь тип `<Runtime>` **RuntimeOverride.**

Атрибут элемента `overrideType` не `<Override>` существует. Разница определяется родительским элементом и типом родительского элемента. Элемент, `<Override>` который находится внутри `<Token>` элемента, который является , должен быть `xsi:type` `RequirementToken` типа **RequirementTokenOverride**. Элемент внутри любого другого родительского элемента или элемента типа должен быть типа `<Override>` `<Override>` `LocaleToken` **LocaleTokenOverride.** Дополнительные сведения об использовании этого элемента, когда он является ребенком элемента, см. в этой ссылке Работа с расширенными `<Token>` [переопределениями манифеста.](../../develop/extended-overrides.md)

Каждый тип описан в отдельных разделах позднее в этой статье.

## <a name="override-element-for-localetoken"></a>Элемент Переопределения для `LocaleToken`

Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление. Если элемент `<Override>` имеет тип **LocaleTokenOverride,** то атрибут является условием, а атрибут `Locale` — `Value` последующим. Например, ниже приводится следующий текст: "Если параметр Office fr-fr, то имя отображения — "Lecteur vidéo".

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

## <a name="override-element-for-requirementtoken"></a>Элемент Переопределения для `RequirementToken`

Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление. Если элемент `<Override>` имеет тип **RequirementTokenOverride,** то детский элемент выражает условие, а атрибут — `<Requirements>` `Value` следовательно. Например, первое из следующих строк гласит: "Если текущая платформа поддерживает `<Override>` версию FeatureOne 1.7, используйте строку "oldAddinVersion" вместо маркера в URL-адресе бабушки и дедушки (вместо строки по умолчанию `${token.requirements}` `<ExtendedOverrides>` "обновление") ".

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
|Значение|string|Обязательный|Значение маркера дедушек и дедушек при условии удовлетворены.|

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

## <a name="override-element-for-runtime"></a>Элемент Переопределения для `Runtime`

> [!IMPORTANT]
> Поддержка этого элемента была представлена в наборе требований к почтовым ящикам [1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) с функцией активации на основе [событий.](../../outlook/autolaunch.md) См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление. Если элемент `<Override>` имеет тип **RuntimeOverride,** то атрибут является условием, а атрибут `type` — `resid` последующим. Например, ниже приводится следующее: "Если тип является "javascript", то это `resid` "JSRuntime.Url". Outlook Этот элемент требуется для обработчиков [точеки расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**Тип надстройки:** почтовая

### <a name="syntax"></a>Синтаксис

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>Содержится в

- [Время выполнения](runtime.md)

### <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|**type**|string|Да|Указывает язык для этого переопределения. В настоящее `"javascript"` время это единственный поддерживаемый вариант.|
|**resid**|string|Да|Указывает расположение URL-адреса файла JavaScript, который должен переопределять расположение URL-адреса HTML по умолчанию, определенного в родительском элементе [Runtime.](runtime.md) `resid` Символ может быть не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента `Resources` элемента.|

### <a name="examples"></a>Примеры

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>См. также

- [Время выполнения](runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md)
