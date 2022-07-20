---
title: Правила активации для надстроек Outlook
description: Outlook активирует некоторые типы надстроек, если сообщение или сведения о встрече, которые читает или создает пользователь, соответствуют правилам активации надстройки.
ms.date: 12/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: af9edf0254156d7bdac13d0553036a614d8c4c39
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889641"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Правила активации контекстных надстроек Outlook

Outlook активирует некоторые типы надстроек, если сообщение или сведения о встрече, которые читает или создает пользователь, соответствуют правилам активации надстройки. Это верно для всех надстроек, для которых используется схема манифеста 1.1. Затем пользователь может выбрать надстройку из пользовательского интерфейса Outlook, чтобы запустить ее для текущего элемента.

На следующем изображении показаны надстройки Outlook, активируемые в области надстройки для сообщения в области чтения.

![Панель приложений с активируемыми приложениями для чтения почты.](../images/read-form-app-bar.png)

## <a name="specify-activation-rules-in-a-manifest"></a>Указание правил активации в манифесте

Чтобы Приложение Outlook активирует надстройку для определенных условий, `Rule` укажите правила активации в манифесте надстройки, используя один из следующих элементов.

- [элемент Rule (MailApp complexType)](/javascript/api/manifest/rule), задающий отдельное правило;
- [Элемент Rule (RuleCollection complexType)](/javascript/api/manifest/rule#rulecollection), совмещающий несколько правил с помощью логических операторов.

 > [!NOTE]
 > Элемент `Rule` , используемый для указания отдельного правила, имеет сложный [тип](/javascript/api/manifest/rule) абстрактного правила. Каждый из следующих типов правил расширяет этот абстрактный `Rule` сложный тип. Следовательно, указывая отдельное правило в манифесте, необходимо использовать атрибут [xsi:type](https://www.w3.org/TR/xmlschema-1/), чтобы определить один из перечисленных ниже типов правил.
 >
 > Например, следующее правило определяет правило [ItemIs](/javascript/api/manifest/rule#itemis-rule) .
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 >
 > Атрибут `FormType` применяется к правилам активации в манифесте версии 1.1, но не определен в `VersionOverrides` версии 1.0. Поэтому его нельзя использовать при использовании [ItemIs](/javascript/api/manifest/rule#itemis-rule) в узле `VersionOverrides` .

В таблице ниже перечислены доступные типы элементов. Дополнительные сведения см. под таблицей и в статьях, перечисленных в статье [Создание надстроек Outlook для форм чтения](read-scenario.md).

|**Имя правила**|**Применимые формы**|**Описание**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|Чтение, создание|Проверяет, относится ли текущий элемент к определенному типу (сообщение или встреча). Кроме того, оно может проверять класс элемента, тип формы и, при необходимости, класс сообщения элемента.|
|[ItemHasAttachment](#itemhasattachment-rule)|Чтение|Проверяет, содержит ли выделенный элемент вложение.|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|Чтение|Проверяет, содержит ли выделенный элемент одну или несколько известных сущностей. Дополнительные сведения см. в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|Чтение|Проверяет, содержит ли адрес электронной почты отправителя, тема и/или тело выбранного элемента совпадение с регулярным выражением. Подробнее: [Использование регулярных правил активации выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](#rulecollection-rule)|Чтение, создание|Объединяет набор правил, чтобы можно было создавать более сложные правила.|

## <a name="itemis-rule"></a>Правило ItemIs

Сложный `ItemIs` тип определяет `true` правило, результатом которого является совпадение текущего элемента с типом элемента и при необходимости класса сообщения элемента, если оно указано в правиле.

Укажите один из следующих типов элементов в атрибуте `ItemType` `ItemIs` правила. В манифесте можно указать `ItemIs` несколько правил. Значение simpleType атрибута ItemType определяет типы элементов Outlook, поддерживающих надстройки Outlook.

|**Value**|**Описание**|
|:-----|:-----|
|**Встреча**|Указывает элемент в календаре Outlook. Это может быть элемент собрания, для которого был отправлен ответ и у которого есть организатор и участники, или встреча без организатора или участника, которая просто представляет собой элемент календаря. Соответствует классу сообщений IPM.Appointment в Outlook.|
|**Сообщение**|Указывает один из следующих элементов, полученных в папке "Входящие". <ul><li><p>Сообщение электронной почты. Соответствует классу сообщений IPM.Note в Outlook.</p></li><li><p>Запрос на собрание, ответ или отклонение. Это соответствует следующим классам сообщений в Outlook.</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

Атрибут `FormType` используется для указания режима (чтения или создания), в котором должна активироваться надстройка.

 > [!NOTE]
 > Атрибут ItemIs `FormType` определяется в схеме версии 1.1 и более поздних версий, но не в `VersionOverrides` версии 1.0. Не включайте атрибут `FormType` при определении команд надстройки.

После активации надстройки можно использовать свойство [mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) для получения элемента, выбранного в текущий момент в Outlook, и свойство [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) для получения типа текущего элемента.

При необходимости `ItemClass` можно использовать атрибут для указания класса сообщения элемента и атрибут, чтобы указать, `IncludeSubClasses` `true` должно ли правило быть, если элемент является подклассом указанного класса.

Дополнительные сведения о классах сообщений см. в статье [Типы элементов и классы сообщений](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).

В следующем примере показано `ItemIs` правило, которое позволяет пользователям просматривать надстройку на панели надстроек Outlook, когда пользователь читает сообщение.

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

В следующем примере показано `ItemIs` правило, которое позволяет пользователям просматривать надстройку на панели надстроек Outlook, когда пользователь читает сообщение или встречу.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```

## <a name="itemhasattachment-rule"></a>Правило ItemHasAttachment

Сложный `ItemHasAttachment` тип определяет правило, которое проверяет, содержит ли выбранный элемент вложение.

```xml
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>Правило ItemHasKnownEntity

Перед тем как элемент становится доступным надстройке, сервер проверяет, содержат ли тема и основной текст строку, которая с высокой вероятностью может быть одной из известных сущностей. Если найдена любая из этих сущностей, `getEntities` `getEntitiesByType` она помещается в коллекцию известных сущностей, доступ к которых можно получить с помощью этого элемента или метода.

Можно указать правило, которое `ItemHasKnownEntity` показывает надстройку, если в элементе присутствует сущность указанного типа. В атрибуте правила можно `EntityType` `ItemHasKnownEntity` указать следующие известные сущности.

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL-адрес

При необходимости можно включить `RegularExpression` регулярное выражение в атрибут, чтобы надстройка отображалася только в том случае, если сущность соответствует регулярному выражению в настоящем. Чтобы получить совпадения с регулярными выражениями `ItemHasKnownEntity` , указанными в правилах, `getRegExMatches` можно использовать метод или метод `getFilteredEntitiesByName` для текущего выбранного элемента Outlook.

В следующем примере показана коллекция `Rule` элементов, которые показывают надстройку, когда в сообщении присутствует одна из указанных известных сущностей.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

В следующем примере `ItemHasKnownEntity` `RegularExpression` показано правило с атрибутом, которое активирует надстройку, если в сообщении присутствует URL-адрес, содержащий слово contoso.

```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

Дополнительные сведения о сущностях в правилах активации см. в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).

## <a name="itemhasregularexpressionmatch-rule"></a>Правило ItemHasRegularExpressionMatch

Сложный `ItemHasRegularExpressionMatch` тип определяет правило, которое использует регулярное выражение для сопоставления содержимого указанного свойства элемента. Если текст, соответствующий регулярному выражению, обнаруживается в заданном свойстве элемента, Outlook активирует панель надстроек и отображает надстройку. Для получения совпадений `getRegExMatches` `getRegExMatchesByName` для указанного регулярного выражения можно использовать объект или метод объекта, представляющего текущий выбранный элемент.

В следующем примере `ItemHasRegularExpressionMatch` показано, как активировать надстройку, если текст выбранного элемента содержит "apple", "вай" или "мука", игнорируя регистр.

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

Дополнительные сведения об использовании правила см `ItemHasRegularExpressionMatch` . в разделе "Использование правил активации регулярных выражений [для отображения надстройки Outlook"](use-regular-expressions-to-show-an-outlook-add-in.md).

## <a name="rulecollection-rule"></a>Правило RuleCollection

Сложный `RuleCollection` тип объединяет несколько правил в одно правило. Можно указать, следует ли объединять правила в коллекции с логическим ИЛИ или логическим И с помощью атрибута `Mode` .

Если указан логический оператор "И", то для отображения надстройки элемент должен соответствовать всем заданным правилам в коллекции. Если выбран логический оператор "ИЛИ", надстройка будет отображаться при наличии элемента, соответствующего любому из заданных правил в коллекции.

Правила можно объединять `RuleCollection` для формирования сложных правил. Следующий пример активирует надстройку, если пользователь просматривает встречу или сообщение, а тема или основной текст сообщения или встречи содержит адрес.

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

Следующий пример активирует надстройку, если пользователь создает сообщение или просматривает встречу, а тема или основной текст встречи содержит адрес.

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```

## <a name="limits-for-rules-and-regular-expressions"></a>Ограничения для правил и регулярных выражений

Чтобы обеспечить удобство использования надстроек Outlook, следует соблюдать рекомендации по активации и использованию API-интерфейсов. В следующей таблице показаны общие ограничения для регулярных выражений и правил, но существуют определенные правила для различных приложений. Дополнительные сведения см. в разделе "Ограничения для [активации и API JavaScript](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) для надстроек Outlook" и "Устранение неполадок с активацией [надстроек Outlook"](troubleshoot-outlook-add-in-activation.md).

|**Элемент надстройки**|**Рекомендации**|
|:-----|:-----|
|Размер манифеста|Не более 256 КБ.|
|Правила|Не более 15 правил.|
|ItemHasKnownEntity|Полнофункциональный клиент Outlook применит правило к первому мегабайту основного текста, но не будет применять его к остальному тексту.|
|Регулярные выражения|Для правил ItemHasKnownEntity или ItemHasRegularExpressionMatch для всех приложений Outlook:<br><ul><li>Задавайте не более 5 регулярных выражений в правилах активации для надстройки Outlook. Если этот предел будет превышен, установить надстройку будет невозможно.</li><li>Задавайте регулярные выражения, ожидаемые результаты которых возвращаются в первых 50 совпадениях с помощью метода <b>getRegExMatches</b>. </li><li>**Важно**! Текст выделяется на основе строк, полученных в результате сопоставления регулярного выражения. Однако выделенные вхождения могут не совпадать с тем, что должно быть результатом фактических утверждений регулярных выражений, `(?!text)`таких как отрицательный просмотр вперед, `(?<=text)`поиск программной части и отрицательный поиск программной части `(?<!text)`. Например, если `under(?!score)` используется регулярное выражение "Like under, under, under score, and underscore", строка "under" выделяется для всех вхождений, а не только для первых двух.</li><li>Укажите регулярные выражения, соответствие которых не превышает ограничения в следующей таблице.<br/><br/><table><tr><th>Ограничение длины для результата, соответствующего регулярному выражению</th><th>Полнофункциональные клиенты Outlook</th><th>Outlook для iOS и Android</th></tr><tr><td>Основной текст элемента в виде простого текста</td><td>1,5 КБ</td><td>3 КБ</td></tr><tr><td>Основной текст элемента в виде HTML-кода</td><td>3 КБ</td><td>3 КБ</td></tr></table>|

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Ограничения активации и API JavaScript для надстроек Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Использование правил активации на основе регулярных выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
