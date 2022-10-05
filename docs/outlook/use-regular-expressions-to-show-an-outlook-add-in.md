---
title: Использование правил активации на основе регулярных выражений для отображения надстройки
description: Узнайте, как использовать правила активации на основе регулярных выражений для контекстных надстроек Outlook.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: ed2fbbfcf7bf55e04f4ec6f225e29fb43ec99639
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467092"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>Использование правил активации на основе регулярных выражений для отображения надстройки Outlook

Вы можете указать правила на основе регулярных выражений для активации [контекстной надстройки](contextual-outlook-add-ins.md) при обнаружении соответствия в определенных полях сообщения. Контекстные надстройки активируются только в режиме чтения. Outlook не активирует контекстные надстройки, когда пользователь создает элемент. Существуют и другие сценарии, в которых Outlook не активирует надстройки, например элементы с цифровой подписью. Дополнительные сведения см. в статье [Правила активации для надстроек Outlook](activation-rules.md).

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

Вы можете указать регулярное выражение в составе правила [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) или [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста надстройки. Правила указываются в точке расширения [DetectedEntity](/javascript/api/manifest/extensionpoint#detectedentity).

Outlook оценивает регулярные выражения на основе правил для интерпретатора JavaScript, используемых браузером на клиентском компьютере. Outlook поддерживает те же специальные знаки, что и все обработчики XML. Они перечислены в следующей таблице. Эти символы можно использовать в регулярном выражении, указав escape-последовательность соответствующего символа, как описано в следующей таблице.

|Знак|Описание|Escape-последовательность|
|:-----|:-----|:-----|
|`"`|Двойная кавычка|`&quot;`|
|`&`|Амперсанд|`&amp;`|
|`'`|Апостроф|`&apos;`|
|`<`|Знак "меньше"|`&lt;`|
|`>`|Знак "больше"|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>Правило ItemHasRegularExpressionMatch

Правило `ItemHasRegularExpressionMatch` позволяет управлять активацией надстройки в зависимости от определенных значений поддерживаемого свойства. Ниже описаны атрибуты правила `ItemHasRegularExpressionMatch`.

|Имя атрибута|Описание|
|:-----|:-----|
|`RegExName`|Указывает имя регулярного выражения, чтобы вы могли сослаться на него в коде надстройки.|
|`RegExValue`|Указывает регулярное выражение, которое будет рассчитано для определения необходимости отображения надстройки.|
|`PropertyName`|Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения — `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` и `Subject`.<br/><br/>Если вы укажете `BodyAsHTML`, Outlook будет применять регулярное выражение, только если текст элемента представлен в формате HTML. В противном случае Outlook возвращает отсутствие совпадений для этого регулярного выражения.<br/><br/>Если вы укажете `BodyAsPlaintext`, Outlook всегда будет применять регулярное выражение для текста элемента.<br/><br/>**Важно:** Если необходимо указать атрибут **Highlight** **\<Rule\>** для элемента, необходимо задать для **атрибута PropertyName** значение `BodyAsPlaintext`. |
|`IgnoreCase`|Указывает, следует ли игнорировать регистр при поиске соответствий регулярному выражению, заданному атрибутом `RegExName`.|
| `Highlight` | Указывает, как клиент должен выделять соответствующий текст. Этот элемент может применяться только к элементам `Rule`, вложенным в элементы `ExtensionPoint`. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.<br/><br/>**Важно:** Чтобы указать **атрибут Highlight** в элементе **\<Rule\>** , необходимо задать для **атрибута PropertyName** значение `BodyAsPlaintext`. |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>Рекомендации по использованию регулярных выражений в правилах

Обратите особое внимание на следующее при использовании регулярных выражений.

- При указании `ItemHasRegularExpressionMatch` правила в тексте элемента регулярное выражение должно дополнительно фильтровать текст и не должно пытаться вернуть весь текст элемента. Использование регулярного выражения, например `.*` попытки получить весь текст элемента, не всегда возвращает ожидаемые результаты.
- Возвращаемый обычный текст может несколько отличаться в зависимости браузера. Если вы используете правило `ItemHasRegularExpressionMatch` с таким значением атрибута `PropertyName`: `BodyAsPlaintext`, проверьте свое регулярное выражение во всех поддерживаемых надстройкой браузерах.

    Так как в разных браузерах основной текст выбранного элемента считывается разными способами, ваше регулярное выражение должно учитывать мелкие различия, которые могут быть возвращены в составе основного текста. Например, в некоторых браузерах, таких как Internet Explorer 9, для получения основного текста элемента используется свойство `innerText` модели DOM, а в других (например, Firefox) — метод `.textContent()`. Кроме того, различные браузеры могут по-разному возвращать разрывы строк (в Internet Explorer — `\r\n`, а в Firefox и Chrome — `\n`). Дополнительные сведения см. в документе [Консорциум W3C: совместимость с моделью DOM (HTML)](https://quirksmode.org/dom/html/).

- Текст элемента в HTML-формате немного отличается для полнофункционального клиента Outlook, Outlook в Интернете и Outlook для мобильных устройств. Будьте внимательны, задавая регулярные выражения.

- В зависимости от клиента Outlook, типа устройства или свойства, к которому применяется регулярное выражение, существуют другие рекомендации и ограничения для каждого клиента, которые следует учитывать при разработке регулярных выражений в качестве правил активации. Дополнительные сведения см. в разделе "Ограничения для [активации и API JavaScript](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) для надстроек Outlook".

### <a name="examples"></a>Примеры

Следующее правило `ItemHasRegularExpressionMatch` активирует надстройку, если SMTP-адрес отправителя содержит строку `@contoso` без учета регистра.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

Ниже приведен другой способ указания того же регулярного выражения с использованием атрибута `IgnoreCase`.

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

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
> Независимо от языкового стандарта, указанного в манифесте, Outlook может извлекать строки сущностей только на английском языке. Только сообщения поддерживают тип `MeetingSuggestion` сущности; встречи не поддерживают это. Вы не можете извлекать сущности из элементов в папке "Отправленные",  `ItemHasKnownEntity` а также использовать правило для активации надстройки для элементов в папке **"** Отправленные".

Правило `ItemHasKnownEntity` поддерживает атрибуты, перечисленные в следующей таблице. Обратите внимание, что указывать регулярное выражение в правиле `ItemHasKnownEntity` необязательно, но при использовании регулярного выражения в качестве фильтра сущности необходимо указывать атрибуты `RegExFilter` и `FilterName`.

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

Совпадения с регулярным выражением можно получить с помощью следующих методов для текущего элемента.

- Метод [getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) возвращает строки текущего элемента, соответствующие всем регулярным выражениям, указанным в правилах `ItemHasRegularExpressionMatch` и `ItemHasKnownEntity` для надстройки.

- [getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) возвращает строки текущего элемента, соответствующие определенному регулярному выражению, указанному в правиле `ItemHasRegularExpressionMatch` надстройки.

- [getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) возвращает полные экземпляры сущностей, которые содержат соответствия определенному регулярному выражению, указанному в правиле `ItemHasKnownEntity` надстройки.

При оценке регулярных выражений соответствия возвращаются в надстройку в массиве. При использовании метода `getRegExMatches` идентификатор этого массива соответствует имени регулярного выражения.

> [!NOTE]
> Outlook не возвращает совпадения в определенном порядке в массиве. Кроме того, не следует предполагать, что совпадения возвращаются в этом массиве в том же порядке, даже если вы выполняете ту же надстройку на каждом из этих клиентов в одном и том же элементе в одном почтовом ящике.

### <a name="examples"></a>Примеры

Ниже приведен пример коллекции правил, содержащей правило `ItemHasRegularExpressionMatch` с регулярным выражением `videoURL`.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

В следующем примере используется метод `getRegExMatches` текущего элемента, чтобы поместить в переменную `videos` результаты предыдущего правила `ItemHasRegularExpressionMatch`.

```js
const videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.

```js
function initDialer()
{
    let myEntities;
    let myString;
    let myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (let i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

Ниже приведен пример правила `ItemHasKnownEntity`, которое указывает сущность `MeetingSuggestion` и регулярное выражение `CampSuggestion`. Outlook активирует надстройку, если обнаруживает, что выбранный элемент содержит приглашение на собрание, а тема или текст содержит термин `WonderCamp`.

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

В следующем примере кода используется метод `getFilteredEntitiesByName` текущего элемента, чтобы поместить в переменную `suggestions` массив обнаруженных приглашений на собрание для предыдущего правила `ItemHasKnownEntity`.

```js
const suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>См. также

- [Надстройка Outlook: номер заказа Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) — контекстная надстройка, которая активируется на основе соответствия регулярному выражению.
- [Создание надстроек Outlook для форм чтения](read-scenario.md)
- [Правила активации для надстроек Outlook](activation-rules.md)
- [Ограничения для активации и API JavaScript для надстроек Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
- [Рекомендации по использованию регулярных выражений в платформе .NET Framework](/dotnet/standard/base-types/best-practices)
