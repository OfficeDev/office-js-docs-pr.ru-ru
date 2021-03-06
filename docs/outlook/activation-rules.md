---
title: Правила активации для надстроек Outlook
description: Outlook активирует некоторые типы надстроек, если сообщение или сведения о встрече, которые читает или создает пользователь, соответствуют правилам активации надстройки.
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: 24f17b7bb3da4665f3f05b23d34ba15bcc4ae729
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349023"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Правила активации контекстных надстроек Outlook

Outlook активирует некоторые типы надстроек, если сообщение или сведения о встрече, которые читает или создает пользователь, соответствуют правилам активации надстройки. Это верно для всех надстроек, для которых используется схема манифеста 1.1. Затем пользователь может выбрать надстройку из пользовательского интерфейса Outlook, чтобы запустить ее для текущего элемента.

На следующем изображении показаны надстройки Outlook, активируемые в области надстройки для сообщения в области чтения.

![Панели приложений с активированными приложениями для чтения почты.](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a>Указание правил активации в манифесте


Чтобы Outlook активировать надстройки для определенных условий, укажите правила активации в манифесте надстройки с помощью одного из `Rule` следующих элементов.

- [элемент Rule (MailApp complexType)](../reference/manifest/rule.md), задающий отдельное правило;
- [Элемент Rule (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection), совмещающий несколько правил с помощью логических операторов.


 > [!NOTE]
 > Элемент, `Rule` который используется для указания отдельного [](../reference/manifest/rule.md) правила, имеет сложный абстрактный тип Правила. Каждый из следующих типов правил расширяет этот абстрактный `Rule` сложный тип. Следовательно, указывая отдельное правило в манифесте, необходимо использовать атрибут [xsi:type](https://www.w3.org/TR/xmlschema-1/), чтобы определить один из перечисленных ниже типов правил.
 > 
 > Например, в следующем правиле определяется правило [ItemIs.](../reference/manifest/rule.md#itemis-rule)
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > Атрибут применяется к правилам активации `FormType` в манифесте v1.1, но не определен `VersionOverrides` в v1.0. Поэтому его нельзя использовать, когда [ItemIs](../reference/manifest/rule.md#itemis-rule) используется в `VersionOverrides` узле.

В таблице ниже перечислены доступные типы элементов. Дополнительные сведения см. под таблицей и в статьях, перечисленных в статье [Создание надстроек Outlook для форм чтения](read-scenario.md).

<br/>

|**Имя правила**|**Применимые формы**|**Описание**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|Чтение, создание|Проверяет, относится ли текущий элемент к определенному типу (сообщение или встреча). Кроме того, оно может проверять класс элемента, тип формы и, при необходимости, класс сообщения элемента.|
|[ItemHasAttachment](#itemhasattachment-rule)|Чтение|Проверяет, содержит ли выделенный элемент вложение.|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|Чтение|Проверяет, содержит ли выделенный элемент одну или несколько известных сущностей. Дополнительные сведения см. в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|Чтение|Проверяет, содержит ли адрес электронной почты отправителя, тема и/или тело выбранного элемента совпадение с регулярным выражением. Подробнее: [Использование регулярных правил активации выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).|
|[RuleCollection](#rulecollection-rule)|Чтение, создание|Объединяет набор правил, чтобы можно было создавать более сложные правила.|

## <a name="itemis-rule"></a>Правило ItemIs

Сложный тип **ItemIs** определяет правило, которое имеет значение **true**, если текущий элемент совпадает с типом элемента и (необязательно) с классом сообщения элемента (если он указан в правиле).

Укажите один из следующих типов элементов в `ItemType` атрибуте правила **ItemIs.** В манифесте можно указать несколько правил **ItemIs**. Значение simpleType атрибута ItemType определяет типы элементов Outlook, поддерживающих надстройки Outlook.

<br/>

|**Value**|**Описание**|
|:-----|:-----|
|**Встреча**|Указывает элемент в календаре Outlook. Это может быть элемент собрания, для которого был отправлен ответ и у которого есть организатор и участники, или встреча без организатора или участника, которая просто представляет собой элемент календаря. Соответствует классу сообщений IPM.Appointment в Outlook.|
|**Сообщение**|Указывает один из следующих элементов, полученных в обычном почтовом ящике. <ul><li><p>Сообщение электронной почты. Соответствует классу сообщений IPM.Note в Outlook.</p></li><li><p>Запрос на собрание, ответ или отклонение. Это соответствует следующим классам сообщений в Outlook.</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

Атрибут используется для указания режима (чтения или составить), в котором должна активироваться `FormType` надстройка.


 > [!NOTE]
 > Атрибут ItemIs определяется в схеме v1.1 и более поздней, но `FormType` не `VersionOverrides` в v1.0. Не включайте атрибут при определении команд `FormType` надстройки.

После активации надстройки можно использовать свойство [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) для получения элемента, выбранного в текущий момент в Outlook, и свойство [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) для получения типа текущего элемента.

Атрибут можно дополнительно использовать для указания класса сообщений элемента, а атрибут указывает, должно ли правило быть верным, если элемент является подклассом `ItemClass` `IncludeSubClasses` указанного класса. 

Дополнительные сведения о классах сообщений см. в статье [Типы элементов и классы сообщений](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).

В следующем примере приводится правило **ItemIs,** которое позволяет пользователям видеть надстройку в Outlook надстройки при чтении сообщения пользователем.

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

В приведенном ниже примере показано правило **ItemIs**, которое отображает надстройку на панели надстройки Outlook, когда пользователь просматривает сообщение или встречу.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a>Правило ItemHasAttachment


Сложный тип определяет правило, которое проверяет, содержит ли выбранный элемент `ItemHasAttachment` вложение.

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a>Правило ItemHasKnownEntity

Перед тем как элемент становится доступным надстройке, сервер проверяет, содержат ли тема и основной текст строку, которая с высокой вероятностью может быть одной из известных сущностей. Если какие-либо из этих сущностями найдены, он помещается в коллекцию известных сущностями, к которые можно получить доступ с помощью или метода `getEntities` `getEntitiesByType` этого элемента.

Вы можете указать правило, используя которое отображает вашу надстройка, когда в элементе присутствует объект `ItemHasKnownEntity` указанного типа. В атрибуте правила можно указать следующие известные `EntityType` `ItemHasKnownEntity` сущности.

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL-адрес

Можно дополнительно включить регулярное выражение в атрибут, чтобы надстройка была показана только в том случае, если объект соответствует обычному выражению `RegularExpression` в настоящее время. Чтобы получить совпадения с регулярными выражениями, указанными в правилах, можно использовать метод или метод для выбранного Outlook `ItemHasKnownEntity` `getRegExMatches` `getFilteredEntitiesByName` элемента.

В следующем примере показана коллекция элементов, отображающих надстройки, когда в сообщении присутствует одно из указанных известных `Rule` сущностями.

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

В следующем примере показано правило с атрибутом, который активирует надстройку, если в сообщении присутствует URL-адрес, содержащий слово `ItemHasKnownEntity` `RegularExpression` "contoso".


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

Дополнительные сведения о сущностях в правилах активации см. в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).


## <a name="itemhasregularexpressionmatch-rule"></a>Правило ItemHasRegularExpressionMatch

Сложный тип определяет правило, которое использует регулярное выражение, чтобы соответствовать содержимому указанного `ItemHasRegularExpressionMatch` свойства элемента. Если текст, соответствующий регулярному выражению, обнаруживается в заданном свойстве элемента, Outlook активирует панель надстроек и отображает надстройку. Для получения совпадений для указанного регулярного выражения можно использовать объект или метод объекта, который представляет выбранный в настоящее время `getRegExMatches` `getRegExMatchesByName` элемент.

В следующем примере показана надстройка, активируемая, когда тело выбранного элемента содержит `ItemHasRegularExpressionMatch` "яблоко", "банан" или "кокос", игнорируя случай.

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

Дополнительные сведения об использовании правила см. в статью Использование правил активации регулярных выражений, чтобы показать Outlook `ItemHasRegularExpressionMatch` [надстройку.](use-regular-expressions-to-show-an-outlook-add-in.md)


## <a name="rulecollection-rule"></a>Правило RuleCollection


Сложный `RuleCollection` тип объединяет несколько правил в одно правило. Можно указать, следует ли сочетать правила в коллекции с логическим или логическим и с помощью `Mode` атрибута.

Если указан логический оператор "И", то для отображения надстройки элемент должен соответствовать всем заданным правилам в коллекции. Если выбран логический оператор "ИЛИ", надстройка будет отображаться при наличии элемента, соответствующего любому из заданных правил в коллекции.

Вы можете объединить `RuleCollection` правила для формирования сложных правил. Следующий пример активирует надстройку, если пользователь просматривает встречу или сообщение, а тема или основной текст сообщения или встречи содержит адрес.

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


Чтобы обеспечить удобство использования надстроек Outlook, следует соблюдать рекомендации по активации и использованию API-интерфейсов. В следующей таблице показаны общие ограничения для регулярных выражений и правил, но существуют определенные правила для различных приложений. Дополнительные сведения см. в дополнительных сведениях о ограничениях для активации и [API JavaScript](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) для Outlook надстройки и устранения неполадок Outlook [надстройки.](troubleshoot-outlook-add-in-activation.md)

<br/>

|**Элемент надстройки**|**Рекомендации**|
|:-----|:-----|
|Размер манифеста|Не более 256 КБ.|
|Правила|Не более 15 правил.|
|ItemHasKnownEntity|Полнофункциональный клиент Outlook применит правило к первому мегабайту основного текста, но не будет применять его к остальному тексту.|
|Регулярные выражения|Правила ItemHasKnownEntity или ItemHasRegularExpressionMatch для всех Outlook приложений:<br><ul><li>Задавайте не более 5 регулярных выражений в правилах активации для надстройки Outlook. Если этот предел будет превышен, установить надстройку будет невозможно.</li><li>Задавайте регулярные выражения, ожидаемые результаты которых возвращаются в первых 50 совпадениях с помощью метода <b>getRegExMatches</b>. </li><li>Указывайте в регулярных выражениях утверждения с просмотром вперед, а не утверждения с просмотром назад `(?<=text)` или отрицательные утверждения с просмотром назад `(?<!text)`.</li><li>Задавайте регулярные выражения, соответствия которым не превышают ограничений, указанных в приведенной ниже таблице.<br/><br/><table><tr><th>Ограничение длины для результата, соответствующего регулярному выражению</th><th>Полнофункциональные клиенты Outlook</th><th>Outlook для iOS и Android</th></tr><tr><td>Основной текст элемента в виде простого текста</td><td>1,5 КБ</td><td>3 КБ</td></tr><tr><td>Основной текст элемента в виде HTML-кода</td><td>3 КБ</td><td>3 КБ</td></tr></table>|

## <a name="see-also"></a>См. также

- [Создание надстроек Outlook для форм создания](compose-scenario.md)
- [Ограничения активации и API JavaScript для надстроек Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Использование правил активации на основе регулярных выражений для отображения надстройки Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md)
    
