# <a name="rule-element"></a><span data-ttu-id="59ed9-101">Элемент Rule</span><span class="sxs-lookup"><span data-stu-id="59ed9-101">Rule element</span></span>

<span data-ttu-id="59ed9-102">Задает правило (или правила) активации, которое следует оценивать для этой контекстной почтовой надстройки.</span><span class="sxs-lookup"><span data-stu-id="59ed9-102">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="59ed9-103">**Тип надстройки:** контекстная почтовая надстройка</span><span class="sxs-lookup"><span data-stu-id="59ed9-103">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="59ed9-104">Содержится в</span><span class="sxs-lookup"><span data-stu-id="59ed9-104">Contained in:</span></span>

- [<span data-ttu-id="59ed9-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="59ed9-105">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="59ed9-106">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="59ed9-106">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="59ed9-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="59ed9-107">Attributes</span></span>

| <span data-ttu-id="59ed9-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="59ed9-108">Attribute</span></span> | <span data-ttu-id="59ed9-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="59ed9-109">Required</span></span> | <span data-ttu-id="59ed9-110">Описание</span><span class="sxs-lookup"><span data-stu-id="59ed9-110">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="59ed9-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="59ed9-111">**xsi:type**</span></span> | <span data-ttu-id="59ed9-112">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-112">Yes</span></span> | <span data-ttu-id="59ed9-113">Тип определяемого правила.</span><span class="sxs-lookup"><span data-stu-id="59ed9-113">The type of rule being defined.</span></span> |

<span data-ttu-id="59ed9-114">Правило может относиться к одному из указанных ниже типов.</span><span class="sxs-lookup"><span data-stu-id="59ed9-114">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="59ed9-115">ItemIs</span><span class="sxs-lookup"><span data-stu-id="59ed9-115">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="59ed9-116">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="59ed9-116">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="59ed9-117">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="59ed9-117">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="59ed9-118">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="59ed9-118">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="59ed9-119">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="59ed9-119">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="59ed9-120">Правило ItemIs</span><span class="sxs-lookup"><span data-stu-id="59ed9-120">ItemIs rule</span></span>

<span data-ttu-id="59ed9-121">Определяет правило, которое оценивается как истинное, если выбранный элемент относится к указанному типу.</span><span class="sxs-lookup"><span data-stu-id="59ed9-121">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="59ed9-122">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="59ed9-122">Attributes</span></span>

| <span data-ttu-id="59ed9-123">Атрибут</span><span class="sxs-lookup"><span data-stu-id="59ed9-123">Attribute</span></span> | <span data-ttu-id="59ed9-124">Обязательный</span><span class="sxs-lookup"><span data-stu-id="59ed9-124">Required</span></span> | <span data-ttu-id="59ed9-125">Описание</span><span class="sxs-lookup"><span data-stu-id="59ed9-125">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="59ed9-126">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="59ed9-126">**ItemType**</span></span> | <span data-ttu-id="59ed9-127">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-127">Yes</span></span> | <span data-ttu-id="59ed9-p101">Задает сопоставляемый тип элемента. Допустимые значения: `Message` и `Appointment`. К типу элементов `Message` относятся электронные письма, приглашения на собрания, ответы на них и уведомления об их отмене.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="59ed9-131">**FormType**</span><span class="sxs-lookup"><span data-stu-id="59ed9-131">**FormType**</span></span> | <span data-ttu-id="59ed9-132">Нет (в [ExtensionPoint](extensionpoint.md)), да (в [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="59ed9-132">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="59ed9-p102">Указывает, должно ли приложение отображаться в форме чтения или редактирования элемента. Допустимые значения: `Read`, `Edit`, `ReadOrEdit`. Для объекта `Rule` в `ExtensionPoint` НЕОБХОДИМО использовать значение `Read`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="59ed9-136">**itemClass**</span><span class="sxs-lookup"><span data-stu-id="59ed9-136">**itemClass**</span></span> | <span data-ttu-id="59ed9-137">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-137">No</span></span> | <span data-ttu-id="59ed9-p103">Указывает сопоставляемый специализированный класс сообщений. Дополнительные сведения см. в статье [Активация почтовой надстройки в Outlook для определенного класса сообщений](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="59ed9-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="59ed9-140">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="59ed9-140">**IncludeSubClasses**</span></span> | <span data-ttu-id="59ed9-141">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-141">No</span></span> | <span data-ttu-id="59ed9-142">Указывает, должно ли правило оцениваться как истинное, если элемент принадлежит к подклассу указанного класса сообщений; по умолчанию используется значение `false`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-142">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="59ed9-143">Пример</span><span class="sxs-lookup"><span data-stu-id="59ed9-143">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="59ed9-144">Правило ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="59ed9-144">ItemHasAttachment rule</span></span>

<span data-ttu-id="59ed9-145">Определяет правило, которое оценивается как истинное, если элемент содержит вложение.</span><span class="sxs-lookup"><span data-stu-id="59ed9-145">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="59ed9-146">Пример</span><span class="sxs-lookup"><span data-stu-id="59ed9-146">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="59ed9-147">Правило ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="59ed9-147">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="59ed9-148">Определяет правило, которое оценивается как истинное, если элемент содержит текст указанного типа сущности в теме или основном тексте.</span><span class="sxs-lookup"><span data-stu-id="59ed9-148">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="59ed9-149">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="59ed9-149">Attributes</span></span>

| <span data-ttu-id="59ed9-150">Атрибут</span><span class="sxs-lookup"><span data-stu-id="59ed9-150">Attribute</span></span> | <span data-ttu-id="59ed9-151">Обязательный</span><span class="sxs-lookup"><span data-stu-id="59ed9-151">Required</span></span> | <span data-ttu-id="59ed9-152">Описание</span><span class="sxs-lookup"><span data-stu-id="59ed9-152">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="59ed9-153">**entityType**</span><span class="sxs-lookup"><span data-stu-id="59ed9-153">**EntityType**</span></span> | <span data-ttu-id="59ed9-154">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-154">Yes</span></span> | <span data-ttu-id="59ed9-p104">Задает тип сущности, который должен быть обнаружен, чтобы правило было оценено как истинное. Допустимые значения: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` и `Contact`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="59ed9-157">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="59ed9-157">**RegExFilter**</span></span> | <span data-ttu-id="59ed9-158">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-158">No</span></span> | <span data-ttu-id="59ed9-159">Задает регулярное выражение, которое должно выполняться в этой сущности для активации.</span><span class="sxs-lookup"><span data-stu-id="59ed9-159">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="59ed9-160">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="59ed9-160">**FilterName**</span></span> | <span data-ttu-id="59ed9-161">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-161">No</span></span> | <span data-ttu-id="59ed9-162">Задает имя фильтра регулярных выражений, чтобы на этот фильтр можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="59ed9-162">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="59ed9-163">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="59ed9-163">**IgnoreCase**</span></span> | <span data-ttu-id="59ed9-164">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-164">No</span></span> | <span data-ttu-id="59ed9-165">Указывает, следует ли игнорировать регистр при оценке регулярного выражения, заданного атрибутом **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="59ed9-165">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="59ed9-166">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="59ed9-166">**highlight**</span></span> | <span data-ttu-id="59ed9-167">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-167">No</span></span> | <span data-ttu-id="59ed9-p105">**Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующие сущности. Допустимые значения: `all` или `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="59ed9-172">Пример</span><span class="sxs-lookup"><span data-stu-id="59ed9-172">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="59ed9-173">Правило ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="59ed9-173">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="59ed9-174">Задает правило, которое оценивается как истинное, если в указанном свойстве элемента обнаруживается соответствие для указанного регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="59ed9-174">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="59ed9-175">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="59ed9-175">Attributes</span></span>

| <span data-ttu-id="59ed9-176">Атрибут</span><span class="sxs-lookup"><span data-stu-id="59ed9-176">Attribute</span></span> | <span data-ttu-id="59ed9-177">Обязательный</span><span class="sxs-lookup"><span data-stu-id="59ed9-177">Required</span></span> | <span data-ttu-id="59ed9-178">Описание</span><span class="sxs-lookup"><span data-stu-id="59ed9-178">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="59ed9-179">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="59ed9-179">**RegExName**</span></span> | <span data-ttu-id="59ed9-180">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-180">Yes</span></span> | <span data-ttu-id="59ed9-181">Указывает имя регулярного выражения, чтобы на него можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="59ed9-181">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="59ed9-182">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="59ed9-182">**RegExValue**</span></span> | <span data-ttu-id="59ed9-183">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-183">Yes</span></span> | <span data-ttu-id="59ed9-184">Указывает регулярное выражение, которое будет вычислено, чтобы определить, требуется ли отображать надстройку.</span><span class="sxs-lookup"><span data-stu-id="59ed9-184">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="59ed9-185">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="59ed9-185">**PropertyName**</span></span> | <span data-ttu-id="59ed9-186">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-186">Yes</span></span> | <span data-ttu-id="59ed9-p106">Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения: `Subject`, `BodyAsPlaintext`, `BodyAsHtml` и `SenderSTMPAddress`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHtml`, or `SenderSTMPAddress`.</span></span> |
| <span data-ttu-id="59ed9-189">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="59ed9-189">**IgnoreCase**</span></span> | <span data-ttu-id="59ed9-190">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-190">No</span></span> | <span data-ttu-id="59ed9-191">Указывает, следует ли игнорировать регистр при выполнении регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="59ed9-191">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="59ed9-192">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="59ed9-192">**highlight**</span></span> | <span data-ttu-id="59ed9-193">Нет</span><span class="sxs-lookup"><span data-stu-id="59ed9-193">No</span></span> | <span data-ttu-id="59ed9-p107">**Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующий текст. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="59ed9-198">Пример</span><span class="sxs-lookup"><span data-stu-id="59ed9-198">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHtml" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="59ed9-199">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="59ed9-199">RuleCollection</span></span>

<span data-ttu-id="59ed9-200">Задает коллекцию правил и логический оператор, который должен использоваться при их оценке.</span><span class="sxs-lookup"><span data-stu-id="59ed9-200">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="59ed9-201">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="59ed9-201">Attributes</span></span>

| <span data-ttu-id="59ed9-202">Атрибут</span><span class="sxs-lookup"><span data-stu-id="59ed9-202">Attribute</span></span> | <span data-ttu-id="59ed9-203">Обязательный</span><span class="sxs-lookup"><span data-stu-id="59ed9-203">Required</span></span> | <span data-ttu-id="59ed9-204">Описание</span><span class="sxs-lookup"><span data-stu-id="59ed9-204">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="59ed9-205">**Режим**</span><span class="sxs-lookup"><span data-stu-id="59ed9-205">**Mode**</span></span> | <span data-ttu-id="59ed9-206">Да</span><span class="sxs-lookup"><span data-stu-id="59ed9-206">Yes</span></span> | <span data-ttu-id="59ed9-p108">Указывает логический оператор, используемый при оценке коллекции правил. Допустимые значения: `And` или `Or`.</span><span class="sxs-lookup"><span data-stu-id="59ed9-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="59ed9-209">Пример</span><span class="sxs-lookup"><span data-stu-id="59ed9-209">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="59ed9-210">См. также</span><span class="sxs-lookup"><span data-stu-id="59ed9-210">See also</span></span>

- [<span data-ttu-id="59ed9-211">Правила активации для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="59ed9-211">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="59ed9-212">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="59ed9-212">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="59ed9-213">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="59ed9-213">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)