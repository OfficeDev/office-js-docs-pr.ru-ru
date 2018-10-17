# <a name="rule-element"></a>Элемент Rule

Задает правило (или правила) активации, которое следует оценивать для этой контекстной почтовой надстройки.

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
| **ItemType** | Да | Задает сопоставляемый тип элемента. Допустимые значения: `Message` и `Appointment`. К типу элементов `Message` относятся электронные письма, приглашения на собрания, ответы на них и уведомления об их отмене. |
| **FormType** | Нет (в [ExtensionPoint](extensionpoint.md)), да (в [OfficeApp](officeapp.md)) | Указывает, должно ли приложение отображаться в форме чтения или редактирования элемента. Допустимые значения: `Read`, `Edit`, `ReadOrEdit`. Для объекта `Rule` в `ExtensionPoint` НЕОБХОДИМО использовать значение `Read`. |
| **itemClass** | Нет | Указывает сопоставляемый специализированный класс сообщений. Дополнительные сведения см. в статье [Активация почтовой надстройки в Outlook для определенного класса сообщений](https://docs.microsoft.com/outlook/add-ins/activation-rules). |
| **IncludeSubClasses** | Нет | Указывает, должно ли правило оцениваться как истинное, если элемент принадлежит к подклассу указанного класса сообщений; по умолчанию используется значение `false`. |

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
| **entityType** | Да | Задает тип сущности, который должен быть обнаружен, чтобы правило было оценено как истинное. Допустимые значения: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` и `Contact`. |
| **RegExFilter** | Нет | Задает регулярное выражение, которое должно выполняться в этой сущности для активации. |
| **FilterName** | Нет | Задает имя фильтра регулярных выражений, чтобы на этот фильтр можно было ссылаться в коде надстройки. |
| **IgnoreCase** | Нет | Указывает, следует ли игнорировать регистр при оценке регулярного выражения, заданного атрибутом **RegExFilter**. |
| **Highlight** | Нет | **Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующие сущности. Допустимые значения: `all` или `none`. Если этот атрибут не задан, по умолчанию используется значение `all`. |

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
| **PropertyName** | Да | Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения: `Subject`, `BodyAsPlaintext`, `BodyAsHtml` и `SenderSTMPAddress`. |
| **IgnoreCase** | Нет | Указывает, следует ли игнорировать регистр при выполнении регулярного выражения. |
| **Highlight** | Нет | **Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующий текст. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`. |

### <a name="example"></a>Пример

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a>RuleCollection

Задает коллекцию правил и логический оператор, который должен использоваться при их оценке.

### <a name="attributes"></a>Атрибуты

| Атрибут | Обязательный | Описание |
|:-----|:-----|:-----|
| **Режим** | Да | Указывает логический оператор, используемый при оценке коллекции правил. Допустимые значения: `And` или `Or`. |

### <a name="example"></a>Пример

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a>См. также

- [Правила активации для надстроек Outlook](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [Сопоставление строк в элементе Outlook как известных сущностей](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [Использование правил активации на основе регулярных выражений для отображения надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)