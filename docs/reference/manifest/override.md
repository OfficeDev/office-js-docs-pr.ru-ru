---
title: Элемент Override в файле манифеста
description: Элемент Override позволяет указать значение параметра в зависимости от заданного состояния.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 131d72883d050038e2df5b7d8bbca033af9e6ee4
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555159"
---
# <a name="override-element"></a>Элемент Override

Предоставляет способ переопределить значение параметра манифеста в зависимости от заданного состояния. Существует три вида условий:

- В Office, который отличается от по `LocaleToken` умолчанию, называется **LocaleTokenOverride**.
- Шаблон поддержки набора требований, который отличается от шаблона по `RequirementToken` умолчанию, **называемого RequirementTokenOverride**.
- Источник отличается от по `Runtime` умолчанию, называется **RuntimeOverride (в настоящее** время в предварительном просмотре).

Элемент, `<Override>` который находится внутри `<Runtime>` элемента, должен быть типа **RuntimeOverride.**

Атрибут элемента `overrideType` не `<Override>` существует. Разница определяется родительским элементом и типом родительского элемента. Элемент, `<Override>` который находится внутри `<Token>` элемента, который `xsi:type` `RequirementToken` является, должен быть типа **RequirementTokenOverride**. Элемент `<Override>` внутри любого другого родительского элемента, или `<Override>` внутри элемента `LocaleToken` типа, должен быть типа **LocaleTokenOverride.** Для получения дополнительной информации об использовании этого элемента, когда он является ребенком `<Token>` элемента, см [Работа с расширенными переопределениями манифеста.](../../develop/extended-overrides.md)

Каждый тип описан в отдельных разделах позже в этой статье.

## <a name="override-element-for-localetoken"></a>Элемент переопределения для `LocaleToken`

Элемент `<Override>` выражает условный и может быть прочитан как "Если ... затем ..." утверждение. Если `<Override>` элемент типа **LocaleTokenOverride**, `Locale` то атрибут является условием, и атрибут является `Value` последующим. Например, ниже приводится следующее: "Если Office настройки является fr-fr, то имя дисплея -" Lecteur vid'o".

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

## <a name="override-element-for-requirementtoken"></a>Элемент переопределения для `RequirementToken`

Элемент `<Override>` выражает условный и может быть прочитан как "Если ... затем ..." утверждение. Если `<Override>` элемент типа **RequirementTokenOverride**, то `<Requirements>` элемент ребенка выражает условие, и атрибут является `Value` последующим. Например, первый из `<Override>` следующих строк читается: "Если текущая платформа поддерживает версию FeatureOne 1.7, то используйте строку 'oldAddinVersion' вместо `${token.requirements}` маркера в URL дедушки и дедушки `<ExtendedOverrides>` (вместо строки по умолчанию 'обновление')".

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
|Значение|string|Обязательный|Значение знака бабушки и дедушки, когда условие удовлетворено.|

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

## <a name="override-element-for-runtime-preview"></a>Элемент переопределения `Runtime` для (предварительный просмотр)

> [!IMPORTANT]
> Эта функция поддерживается только для [предварительного](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) просмотра Outlook веб-сайтах и Windows с Microsoft 365 подпиской. Для получения более подробной [информации см Outlook.](../../outlook/autolaunch.md)
>
> Поскольку функции предварительного просмотра могут быть изменения без предварительного уведомления, они не должны использоваться в производственных дополнениях.

Элемент `<Override>` выражает условный и может быть прочитан как "Если ... затем ..." утверждение. Если `<Override>` элемент типа **RuntimeOverride**, то `type` атрибут является условием, и `resid` атрибут является последующим. Например, ниже приводится следующее: "Если тип "JavaScript", `resid` то 'JSRuntime.Url'." Outlook Рабочий стол требует этого элемента [для обработчиков токов точки расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)

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
|**type**|string|Да|Определяет язык для этого переопределения. В настоящее `"javascript"` время это единственный поддерживаемый вариант.|
|**resid**|string|Да|Определяется местоположение URL-адреса файла JavaScript, который должен переопределить местоположение URL HTML по умолчанию, определяемого в [родительском](runtime.md) элементе `resid` Runtime. Может `resid` быть не более 32 символов и должен `id` соответствовать атрибуту `Url` элемента в `Resources` элементе.|

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
- [Настройте Outlook для активации на основе событий](../../outlook/autolaunch.md)
