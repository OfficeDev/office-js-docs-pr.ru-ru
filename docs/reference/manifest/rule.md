---
title: Элемент Rule в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 78fb38d8fb18c276bfe2eed1bd5b52659cadcaa3
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165533"
---
# <a name="rule-element"></a>Элемент Rule

Указывает правила активации, которые следует оценивать для этой контекстной почтовой надстройки.

**Тип надстройки:** контекстная почтовая надстройка

## <a name="contained-in"></a>Содержится в

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md)

## <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **xsi:type** | Да | Тип определяемого правила. |

Правило может относиться к одному из указанных ниже типов.

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## <a name="itemis-rule"></a>Правило ItemIs

Определяет правило, которое оценивается как истинное, если выбранный элемент относится к указанному типу.

### <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **ItemType** | Да | Указывает сопоставляемый тип элемента. Допустимые значения: `Message` и `Appointment`. К типу элементов `Message` относятся электронные письма, приглашения на собрания, ответы на них и уведомления об их отмене. |
| **FormType** | Нет (в [ExtensionPoint](extensionpoint.md)), да (в [OfficeApp](officeapp.md)) | Указывает, должно ли приложение отображаться в форме чтения или редактирования элемента. Допустимые значения: `Read`, `Edit`, `ReadOrEdit`. Для объекта `Rule` в `ExtensionPoint` НЕОБХОДИМО использовать значение `Read`. |
| **ItemClass** | Нет | Указывает сопоставляемый специализированный класс сообщений. Дополнительные сведения см. в статье [Активация почтовой надстройки в Outlook для определенного класса сообщений](../../outlook/activation-rules.md). |
| **IncludeSubClasses** | Нет | Указывает, должно ли правило оцениваться как истинное (true), если элемент принадлежит к подклассу указанного класса сообщений; по умолчанию используется значение `false`. |

### <a name="example"></a>Пример

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a>Правило ItemHasAttachment

Определяет правило, которое оценивается как истинное, если элемент содержит вложение.

### <a name="example"></a>Пример

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>Правило ItemHasKnownEntity

Определяет правило, которое оценивается как истинное, если элемент содержит текст указанного типа сущности в теме или основном тексте.

### <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **EntityType** | Да | Указывает тип сущности, который должен обнаруживаться, чтобы правило было оценено как истинное. Допустимые значения: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` и `Contact`. |
| **RegExFilter** | Нет | Задает регулярное выражение, которое должно выполняться в этой сущности для активации. |
| **FilterName** | Нет | Задает имя фильтра регулярных выражений, чтобы на этот фильтр можно было ссылаться в коде надстройки. |
| **IgnoreCase** | Нет | Указывает, следует ли игнорировать регистр при сравнении регулярного выражения, заданного атрибутом **RegExFilter**. |
| **Highlight** | Нет | **Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующие сущности. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`. |

### <a name="example"></a>Пример

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a>Правило ItemHasRegularExpressionMatch

Задает правило, которое оценивается как истинное, если в указанном свойстве элемента обнаруживается соответствие для указанного регулярного выражения.

### <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **RegExName** | Да | Указывает имя регулярного выражения, чтобы на него можно было ссылаться в коде надстройки. |
| **RegExValue** | Да | Указывает регулярное выражение, которое будет вычислено, чтобы определить, требуется ли отображать надстройку. |
| **PropertyName** | Да | Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` и `SenderSMTPAddress`.<br/><br/>Если вы укажете `BodyAsHTML`, Outlook будет применять регулярное выражение, только если текст элемента представлен в формате HTML. В противном случае Outlook возвращает отсутствие совпадений для этого регулярного выражения.<br/><br/>Если вы укажете `BodyAsPlaintext`, Outlook всегда будет применять регулярное выражение для текста элемента.<br/><br/>**Примечание.** Необходимо задать атрибут **PropertyName** для `BodyAsPlaintext`, если указан атрибут **Highlight** для элемента **Rule**.|
| **IgnoreCase** | Нет | Указывает, следует ли игнорировать регистр при сравнении регулярного выражения, заданного атрибутом **RegExName**. |
| **Highlight** | Нет | Указывает, как клиент должен выделять соответствующий текст. Этот атрибут может применяться только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.<br/><br/>**Примечание.** Необходимо задать атрибут **PropertyName** для `BodyAsPlaintext`, если указан атрибут **Highlight** для элемента **Rule**.
|

### <a name="example"></a>Пример

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

Задает коллекцию правил и логический оператор, который должен использоваться при их оценке.

### <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **Mode** | Да | Указывает логический оператор, используемый при оценке коллекции правил. Допустимые значения: `And` и `Or`. |

### <a name="example"></a>Пример

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>См. также

- [Правила активации для надстроек Outlook](../../outlook/activation-rules.md)
- [Сопоставление строк в элементе Outlook как известных сущностей](../../outlook/match-strings-in-an-item-as-well-known-entities.md)    
- [Использование регулярных правил активации выражений для отображения надстройки Outlook](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
