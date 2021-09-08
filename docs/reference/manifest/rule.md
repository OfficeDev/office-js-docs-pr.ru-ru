---
title: Элемент Rule в файле манифеста
description: Элемент Правила указывает правила активации, которые должны быть оценены для этой контекстной надстройки почты.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 60882a5e36a63832cf81eab9320b113a420b84a3
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938046"
---
# <a name="rule-element"></a>Элемент Rule

Указывает правила активации, которые необходимо оценить для этой контекстной надстройки почты.

**Тип надстройки:** Почта (контекстная)

## <a name="contained-in"></a>Содержится в

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md) [**(CustomPane** (амортизации)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))

## <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **xsi:type** | Да | Тип определяемого правила. |

Тип правила может быть одним из следующих:

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
