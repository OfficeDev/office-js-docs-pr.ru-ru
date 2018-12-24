---
title: Элемент Rule в файле манифеста
description: ''
ms.date: 11/30/2018
ms.openlocfilehash: ce7763ecb4ef81587ccacbd4090a6f412baf99b2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433118"
---
# <a name="rule-element"></a><span data-ttu-id="c774a-102">Элемент Rule</span><span class="sxs-lookup"><span data-stu-id="c774a-102">Rule element</span></span>

<span data-ttu-id="c774a-103">Указывает правила активации, которые следует оценивать для этой контекстной почтовой надстройки.</span><span class="sxs-lookup"><span data-stu-id="c774a-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="c774a-104">**Тип надстройки:** контекстная почтовая надстройка</span><span class="sxs-lookup"><span data-stu-id="c774a-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="c774a-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c774a-105">Contained in</span></span>

- [<span data-ttu-id="c774a-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c774a-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="c774a-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c774a-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="c774a-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c774a-108">Attributes</span></span>

| <span data-ttu-id="c774a-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c774a-109">Attribute</span></span> | <span data-ttu-id="c774a-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c774a-110">Required</span></span> | <span data-ttu-id="c774a-111">Описание</span><span class="sxs-lookup"><span data-stu-id="c774a-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c774a-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c774a-112">**xsi:type**</span></span> | <span data-ttu-id="c774a-113">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-113">Yes</span></span> | <span data-ttu-id="c774a-114">Тип определяемого правила.</span><span class="sxs-lookup"><span data-stu-id="c774a-114">The type of rule being defined.</span></span> |

<span data-ttu-id="c774a-115">Правило может относиться к одному из указанных ниже типов.</span><span class="sxs-lookup"><span data-stu-id="c774a-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="c774a-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="c774a-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="c774a-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="c774a-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="c774a-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="c774a-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="c774a-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="c774a-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="c774a-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="c774a-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="c774a-121">Правило ItemIs</span><span class="sxs-lookup"><span data-stu-id="c774a-121">ItemIs rule</span></span>

<span data-ttu-id="c774a-122">Определяет правило, которое оценивается как истинное, если выбранный элемент относится к указанному типу.</span><span class="sxs-lookup"><span data-stu-id="c774a-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="c774a-123">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c774a-123">Attributes</span></span>

| <span data-ttu-id="c774a-124">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c774a-124">Attribute</span></span> | <span data-ttu-id="c774a-125">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c774a-125">Required</span></span> | <span data-ttu-id="c774a-126">Описание</span><span class="sxs-lookup"><span data-stu-id="c774a-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c774a-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="c774a-127">**ItemType**</span></span> | <span data-ttu-id="c774a-128">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-128">Yes</span></span> | <span data-ttu-id="c774a-p101">Указывает сопоставляемый тип элемента. Допустимые значения: `Message` и `Appointment`. К типу элементов `Message` относятся электронные письма, приглашения на собрания, ответы на них и уведомления об их отмене.</span><span class="sxs-lookup"><span data-stu-id="c774a-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="c774a-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="c774a-132">**FormType**</span></span> | <span data-ttu-id="c774a-133">Нет (в [ExtensionPoint](extensionpoint.md)), да (в [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="c774a-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="c774a-p102">Указывает, должно ли приложение отображаться в форме чтения или редактирования элемента. Допустимые значения: `Read`, `Edit`, `ReadOrEdit`. Для объекта `Rule` в `ExtensionPoint` НЕОБХОДИМО использовать значение `Read`.</span><span class="sxs-lookup"><span data-stu-id="c774a-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="c774a-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="c774a-137">**ItemClass**</span></span> | <span data-ttu-id="c774a-138">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-138">No</span></span> | <span data-ttu-id="c774a-p103">Указывает сопоставляемый специализированный класс сообщений. Дополнительные сведения см. в статье [Активация почтовой надстройки в Outlook для определенного класса сообщений](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="c774a-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="c774a-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="c774a-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="c774a-142">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-142">No</span></span> | <span data-ttu-id="c774a-143">Указывает, должно ли правило оцениваться как истинное (true), если элемент принадлежит к подклассу указанного класса сообщений; по умолчанию используется значение `false`.</span><span class="sxs-lookup"><span data-stu-id="c774a-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="c774a-144">Пример</span><span class="sxs-lookup"><span data-stu-id="c774a-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="c774a-145">Правило ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="c774a-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="c774a-146">Определяет правило, которое оценивается как истинное, если элемент содержит вложение.</span><span class="sxs-lookup"><span data-stu-id="c774a-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="c774a-147">Пример</span><span class="sxs-lookup"><span data-stu-id="c774a-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="c774a-148">Правило ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="c774a-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="c774a-149">Определяет правило, которое оценивается как истинное, если элемент содержит текст указанного типа сущности в теме или основном тексте.</span><span class="sxs-lookup"><span data-stu-id="c774a-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="c774a-150">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c774a-150">Attributes</span></span>

| <span data-ttu-id="c774a-151">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c774a-151">Attribute</span></span> | <span data-ttu-id="c774a-152">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c774a-152">Required</span></span> | <span data-ttu-id="c774a-153">Описание</span><span class="sxs-lookup"><span data-stu-id="c774a-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c774a-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="c774a-154">**EntityType**</span></span> | <span data-ttu-id="c774a-155">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-155">Yes</span></span> | <span data-ttu-id="c774a-p104">Указывает тип сущности, который должен обнаруживаться, чтобы правило было оценено как истинное. Допустимые значения: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` и `Contact`.</span><span class="sxs-lookup"><span data-stu-id="c774a-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="c774a-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="c774a-158">**RegExFilter**</span></span> | <span data-ttu-id="c774a-159">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-159">No</span></span> | <span data-ttu-id="c774a-160">Задает регулярное выражение, которое должно выполняться в этой сущности для активации.</span><span class="sxs-lookup"><span data-stu-id="c774a-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="c774a-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="c774a-161">**FilterName**</span></span> | <span data-ttu-id="c774a-162">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-162">No</span></span> | <span data-ttu-id="c774a-163">Задает имя фильтра регулярных выражений, чтобы на этот фильтр можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="c774a-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="c774a-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="c774a-164">**IgnoreCase**</span></span> | <span data-ttu-id="c774a-165">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-165">No</span></span> | <span data-ttu-id="c774a-166">Указывает, следует ли игнорировать регистр при оценке регулярного выражения, заданного атрибутом **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="c774a-166">Specifies to ignore case when running the regular expression specified by the  **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="c774a-167">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="c774a-167">**Highlight**</span></span> | <span data-ttu-id="c774a-168">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-168">No</span></span> | <span data-ttu-id="c774a-p105">**Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующие сущности. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="c774a-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="c774a-173">Пример</span><span class="sxs-lookup"><span data-stu-id="c774a-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="c774a-174">Правило ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="c774a-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="c774a-175">Задает правило, которое оценивается как истинное, если в указанном свойстве элемента обнаруживается соответствие для указанного регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="c774a-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="c774a-176">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c774a-176">Attributes</span></span>

| <span data-ttu-id="c774a-177">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c774a-177">Attribute</span></span> | <span data-ttu-id="c774a-178">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c774a-178">Required</span></span> | <span data-ttu-id="c774a-179">Описание</span><span class="sxs-lookup"><span data-stu-id="c774a-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c774a-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="c774a-180">**RegExName**</span></span> | <span data-ttu-id="c774a-181">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-181">Yes</span></span> | <span data-ttu-id="c774a-182">Указывает имя регулярного выражения, чтобы на него можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="c774a-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="c774a-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="c774a-183">**RegExValue**</span></span> | <span data-ttu-id="c774a-184">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-184">Yes</span></span> | <span data-ttu-id="c774a-185">Указывает регулярное выражение, которое будет вычислено, чтобы определить, требуется ли отображать надстройку.</span><span class="sxs-lookup"><span data-stu-id="c774a-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="c774a-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="c774a-186">**PropertyName**</span></span> | <span data-ttu-id="c774a-187">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-187">Yes</span></span> | <span data-ttu-id="c774a-p106">Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` и `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="c774a-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span> |
| <span data-ttu-id="c774a-190">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="c774a-190">**IgnoreCase**</span></span> | <span data-ttu-id="c774a-191">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-191">No</span></span> | <span data-ttu-id="c774a-192">Указывает, следует ли игнорировать регистр при выполнении регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="c774a-192">Specifies to ignore the case when executing the regular expression.</span></span> |
| <span data-ttu-id="c774a-193">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="c774a-193">**Highlight**</span></span> | <span data-ttu-id="c774a-194">Нет</span><span class="sxs-lookup"><span data-stu-id="c774a-194">No</span></span> | <span data-ttu-id="c774a-p107">**Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующий текст. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="c774a-p107">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching text. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="c774a-199">Пример</span><span class="sxs-lookup"><span data-stu-id="c774a-199">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="c774a-200">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="c774a-200">RuleCollection</span></span>

<span data-ttu-id="c774a-201">Задает коллекцию правил и логический оператор, который должен использоваться при их оценке.</span><span class="sxs-lookup"><span data-stu-id="c774a-201">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="c774a-202">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c774a-202">Attributes</span></span>

| <span data-ttu-id="c774a-203">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c774a-203">Attribute</span></span> | <span data-ttu-id="c774a-204">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c774a-204">Required</span></span> | <span data-ttu-id="c774a-205">Описание</span><span class="sxs-lookup"><span data-stu-id="c774a-205">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c774a-206">**Mode**</span><span class="sxs-lookup"><span data-stu-id="c774a-206">**Mode**</span></span> | <span data-ttu-id="c774a-207">Да</span><span class="sxs-lookup"><span data-stu-id="c774a-207">Yes</span></span> | <span data-ttu-id="c774a-p108">Указывает логический оператор, используемый при оценке коллекции правил. Допустимые значения: `And` и `Or`.</span><span class="sxs-lookup"><span data-stu-id="c774a-p108">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="c774a-210">Пример</span><span class="sxs-lookup"><span data-stu-id="c774a-210">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="c774a-211">См. также</span><span class="sxs-lookup"><span data-stu-id="c774a-211">See also</span></span>

- [<span data-ttu-id="c774a-212">Правила активации для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="c774a-212">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="c774a-213">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="c774a-213">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="c774a-214">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="c774a-214">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)