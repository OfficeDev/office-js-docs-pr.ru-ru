---
title: Использование правил активации на основе регулярных выражений для отображения надстройки
description: Узнайте, как использовать правила активации на основе регулярных выражений для контекстных надстроек Outlook.
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: d334ba6b2e0f044fc8d876cd6edd218743ccb390
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936228"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>Использование правил активации на основе регулярных выражений для отображения надстройки Outlook

Вы можете указать правила на основе регулярных выражений для активации [контекстной надстройки](contextual-outlook-add-ins.md) при обнаружении соответствия в определенных полях сообщения. Контекстные надстройки активируются только в режиме чтения. Outlook не активирует контекстные надстройки, когда пользователь создает элемент. Существуют также другие сценарии, в которых Outlook не активирует надстройки, например, элементы с цифровыми подписями. Дополнительные сведения см. в статье [Правила активации для надстроек Outlook](activation-rules.md).

Вы можете указать регулярное выражение в составе правила [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) или [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) в XML-файле манифеста надстройки. Правила указываются в точке расширения [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity).

Outlook оценивает регулярные выражения на основе правил для интерпретатора JavaScript, используемых браузером на клиентском компьютере. Outlook поддерживает те же специальные знаки, что и все обработчики XML. Они перечислены в следующей таблице. Указывая эти знаки в регулярных выражениях, используйте соответствующие escape-последовательности из следующей таблицы.

<br/>

|Знак|Описание|Escape-последовательность|
|:-----|:-----|:-----|
|`"`|Двойная кавычка|`&quot;`|
|`&`|Амперсанд|`&amp;`|
|`'`|Апостроф|`&apos;`|
|`<`|Знак "меньше"|`&lt;`|
|`>`|Знак "больше"|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>Правило ItemHasRegularExpressionMatch

Правило `ItemHasRegularExpressionMatch` позволяет управлять активацией надстройки в зависимости от определенных значений поддерживаемого свойства. Ниже описаны атрибуты правила `ItemHasRegularExpressionMatch`.

<br/>

|Имя атрибута|Описание|
|:-----|:-----|
|`RegExName`|Указывает имя регулярного выражения, чтобы вы могли сослаться на него в коде надстройки.|
|`RegExValue`|Указывает регулярное выражение, которое будет рассчитано для определения необходимости отображения надстройки.|
|`PropertyName`|Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения — `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` и `Subject`.<br/><br/>Если вы укажете `BodyAsHTML`, Outlook будет применять регулярное выражение, только если текст элемента представлен в формате HTML. В противном случае Outlook возвращает отсутствие совпадений для этого регулярного выражения.<br/><br/>Если вы укажете `BodyAsPlaintext`, Outlook всегда будет применять регулярное выражение для текста элемента.<br/><br/>**Примечание.** Необходимо задать атрибут `PropertyName` для `BodyAsPlaintext`, если указан атрибут `Highlight` для элемента `Rule`.|
|`IgnoreCase`|Указывает, следует ли игнорировать регистр при поиске соответствий регулярному выражению, заданному атрибутом `RegExName`.|
| `Highlight` | Указывает, как клиент должен выделять соответствующий текст. Этот элемент может применяться только к элементам `Rule`, вложенным в элементы `ExtensionPoint`. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.<br/><br/>**Примечание.** Необходимо задать атрибут `PropertyName` для `BodyAsPlaintext`, если указан атрибут `Highlight` для элемента `Rule`. |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>Рекомендации по использованию регулярных выражений в правилах

Обратите особое внимание на следующее при использовании регулярных выражений.

- Если вы указываете правило `ItemHasRegularExpressionMatch` для текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.
- Возвращаемый обычный текст может несколько отличаться в зависимости браузера. Если вы используете правило `ItemHasRegularExpressionMatch` с таким значением атрибута `PropertyName`: `BodyAsPlaintext`, проверьте свое регулярное выражение во всех поддерживаемых надстройкой браузерах.

    Так как в разных браузерах основной текст выбранного элемента считывается разными способами, ваше регулярное выражение должно учитывать мелкие различия, которые могут быть возвращены в составе основного текста. Например, в некоторых браузерах, таких как Internet Explorer 9, для получения основного текста элемента используется свойство `innerText` модели DOM, а в других (например, Firefox) — метод `.textContent()`. Кроме того, различные браузеры могут по-разному возвращать разрывы строк (в Internet Explorer — `\r\n`, а в Firefox и Chrome — `\n`). Дополнительные сведения см. в документе [Консорциум W3C: совместимость с моделью DOM (HTML)](https://quirksmode.org/dom/html/).

- Текст элемента в HTML-формате немного отличается для полнофункционального клиента Outlook, Outlook в Интернете и Outlook для мобильных устройств. Будьте внимательны, задавая регулярные выражения.

- В зависимости Outlook клиента, типа устройства или свойства, на которое применяется регулярное выражение, для каждого из клиентов существуют другие методы и ограничения, которые следует знать при разработке регулярных выражений в качестве правил активации. Дополнительные сведения см. в материале Ограничения для активации и [API JavaScript для Outlook надстройки.](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)

### <a name="examples"></a>Примеры

Следующее правило `ItemHasRegularExpressionMatch` активирует надстройку, если SMTP-адрес отправителя содержит строку `@contoso` без учета регистра.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

Ниже приведен другой способ указания того же регулярного выражения с использованием атрибута `IgnoreCase`.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

Следующее правило `ItemHasRegularExpressionMatch` активирует надстройку, если основной текст текущего элемента содержит биржевой символ акции.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a>Правило ItemHasKnownEntity

Правило `ItemHasKnownEntity` активирует надстройку при наличии сущности в теме или тексте выбранного элемента. Тип [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) определяет поддерживаемые сущности. Применять регулярное выражение в правиле `ItemHasKnownEntity` удобно, когда активация надстройки зависит от группы значений сущности (например, определенного набора URL-адресов или номеров телефонов с определенным кодом области).

> [!NOTE]
> Независимо от языкового стандарта, указанного в манифесте, Outlook может извлекать строки сущностей только на английском языке. Только сообщения поддерживают тип сущности `MeetingSuggestion`. Сущности невозможно извлечь из элементов в папке **Отправленные**. Правило `ItemHasKnownEntity` не подходит для активации надстройки для элементов в папке **Отправленные**.

Правило `ItemHasKnownEntity` поддерживает атрибуты, перечисленные в следующей таблице. Обратите внимание, что указывать регулярное выражение в правиле `ItemHasKnownEntity` необязательно, но при использовании регулярного выражения в качестве фильтра сущности необходимо указывать атрибуты `RegExFilter` и `FilterName`.

<br/>

|Имя атрибута|Описание|
|:-----|:-----|
|`EntityType`|Задает тип сущности, который должен быть обнаружен, чтобы правило было оценено как `true`. Используйте несколько правил, чтобы указать несколько типов сущностей.|
|`RegExFilter`|Указывает регулярное выражение, обеспечивающее дальнейшую фильтрацию экземпляров сущности, указанной атрибутом `EntityType`.|
|`FilterName`|Указывает имя регулярного выражения, заданного атрибутом `RegExFilter`, чтобы впоследствии можно было сослаться на него в коде.|
|`IgnoreCase`|Указывает, следует ли игнорировать регистр при поиске соответствий регулярному выражению, заданному атрибутом `RegExFilter`.|

### <a name="examples"></a>Примеры

В следующем правиле `ItemHasKnownEntity` активация надстройки выполняется при наличии URL-адреса в теме или основном тексте текущего элемента и строки `youtube` в этом адресе независимо от регистра.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a>Использование результатов регулярных выражений в коде

Вы можете получить совпадения с регулярным выражением с помощью следующих методов текущего элемента.

- Метод [getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) возвращает строки текущего элемента, соответствующие всем регулярным выражениям, указанным в правилах `ItemHasRegularExpressionMatch` и `ItemHasKnownEntity` для надстройки.

- [getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) возвращает строки текущего элемента, соответствующие определенному регулярному выражению, указанному в правиле `ItemHasRegularExpressionMatch` надстройки.

- [getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) возвращает полные экземпляры сущностей, которые содержат соответствия определенному регулярному выражению, указанному в правиле `ItemHasKnownEntity` надстройки.

При оценке регулярных выражений соответствия возвращаются в надстройку в массиве. При использовании метода `getRegExMatches` идентификатор этого массива соответствует имени регулярного выражения.

> [!NOTE]
> Outlook не возвращает соответствия в каком-либо определенном порядке в массиве. Кроме того, соответствия могут возвращаться в другом порядке, даже если вы запустите ту же настройку в каждом из этих клиентов для того же элемента в том же почтовом ящике.

### <a name="examples"></a>Примеры

Ниже приведен пример коллекции правил, содержащей правило `ItemHasRegularExpressionMatch` с регулярным выражением `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

В следующем примере используется метод `getRegExMatches` текущего элемента, чтобы поместить в переменную `videos` результаты предыдущего правила `ItemHasRegularExpressionMatch`.

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

Несколько совпадений хранятся в этом объекте в виде элементов массива. Следующий пример кода показывает, как выполнять итерацию по совпадениям для регулярного выражения `reg1`, чтобы создать строку для отображения в виде HTML-кода.

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

Ниже приведен пример правила `ItemHasKnownEntity`, которое указывает сущность `MeetingSuggestion` и регулярное выражение `CampSuggestion`. Outlook активирует надстройку, если обнаруживает, что выбранный элемент содержит приглашение на собрание, а тема или текст содержит термин `WonderCamp`.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

В следующем примере кода используется метод `getFilteredEntitiesByName` текущего элемента, чтобы поместить в переменную `suggestions` массив обнаруженных приглашений на собрание для предыдущего правила `ItemHasKnownEntity`.

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>См. также

- [Надстройка Outlook: номер заказа Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) — контекстная надстройка, которая активируется на основе соответствия регулярному выражению.
- [Создание надстроек Outlook для форм чтения](read-scenario.md)
- [Правила активации для надстроек Outlook](activation-rules.md)
- [Ограничения для активации и API JavaScript для надстроек Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
- [Рекомендации по использованию регулярных выражений в .NET Framework](/dotnet/standard/base-types/best-practices)
