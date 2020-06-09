---
title: Элемент Rule в файле манифеста
description: Элемент Rule указывает правила активации, которые должны оцениваться для этой контекстной почтовой надстройки.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: c4094cdf9e9006bbc49d180cb79845527461a543
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608113"
---
# <a name="rule-element"></a><span data-ttu-id="00dde-103">Элемент Rule</span><span class="sxs-lookup"><span data-stu-id="00dde-103">Rule element</span></span>

<span data-ttu-id="00dde-104">Задает правила активации, которые должны оцениваться для этой контекстной почтовой надстройки.</span><span class="sxs-lookup"><span data-stu-id="00dde-104">Specifies the activation rules that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="00dde-105">**Тип надстройки:** Почта (контекстная)</span><span class="sxs-lookup"><span data-stu-id="00dde-105">**Add-in type:** Mail (contextual)</span></span>

## <a name="contained-in"></a><span data-ttu-id="00dde-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="00dde-106">Contained in</span></span>

- [<span data-ttu-id="00dde-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="00dde-107">OfficeApp</span></span>](officeapp.md)
- <span data-ttu-id="00dde-108">[ExtensionPoint](extensionpoint.md) ([**кустомпане** (устаревшее)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span><span class="sxs-lookup"><span data-stu-id="00dde-108">[ExtensionPoint](extensionpoint.md) ([**CustomPane** (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), [**DetectedEntity**](extensionpoint.md#detectedentity))</span></span>

## <a name="attributes"></a><span data-ttu-id="00dde-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00dde-109">Attributes</span></span>

| <span data-ttu-id="00dde-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="00dde-110">Attribute</span></span> | <span data-ttu-id="00dde-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="00dde-111">Required</span></span> | <span data-ttu-id="00dde-112">Описание</span><span class="sxs-lookup"><span data-stu-id="00dde-112">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="00dde-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="00dde-113">**xsi:type**</span></span> | <span data-ttu-id="00dde-114">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-114">Yes</span></span> | <span data-ttu-id="00dde-115">Тип определяемого правила.</span><span class="sxs-lookup"><span data-stu-id="00dde-115">The type of rule being defined.</span></span> |

<span data-ttu-id="00dde-116">Правило может относиться к одному из указанных ниже типов.</span><span class="sxs-lookup"><span data-stu-id="00dde-116">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="00dde-117">ItemIs</span><span class="sxs-lookup"><span data-stu-id="00dde-117">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="00dde-118">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="00dde-118">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="00dde-119">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="00dde-119">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="00dde-120">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="00dde-120">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="00dde-121">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="00dde-121">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="00dde-122">Правило ItemIs</span><span class="sxs-lookup"><span data-stu-id="00dde-122">ItemIs rule</span></span>

<span data-ttu-id="00dde-123">Определяет правило, которое оценивается как истинное, если выбранный элемент относится к указанному типу.</span><span class="sxs-lookup"><span data-stu-id="00dde-123">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="00dde-124">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00dde-124">Attributes</span></span>

| <span data-ttu-id="00dde-125">Атрибут</span><span class="sxs-lookup"><span data-stu-id="00dde-125">Attribute</span></span> | <span data-ttu-id="00dde-126">Обязательный</span><span class="sxs-lookup"><span data-stu-id="00dde-126">Required</span></span> | <span data-ttu-id="00dde-127">Описание</span><span class="sxs-lookup"><span data-stu-id="00dde-127">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="00dde-128">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="00dde-128">**ItemType**</span></span> | <span data-ttu-id="00dde-129">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-129">Yes</span></span> | <span data-ttu-id="00dde-p101">Указывает сопоставляемый тип элемента. Допустимые значения: `Message` и `Appointment`. К типу элементов `Message` относятся электронные письма, приглашения на собрания, ответы на них и уведомления об их отмене.</span><span class="sxs-lookup"><span data-stu-id="00dde-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="00dde-133">**FormType**</span><span class="sxs-lookup"><span data-stu-id="00dde-133">**FormType**</span></span> | <span data-ttu-id="00dde-134">Нет (в [ExtensionPoint](extensionpoint.md)), да (в [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="00dde-134">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="00dde-p102">Указывает, должно ли приложение отображаться в форме чтения или редактирования элемента. Допустимые значения: `Read`, `Edit`, `ReadOrEdit`. Для объекта `Rule` в `ExtensionPoint` НЕОБХОДИМО использовать значение `Read`.</span><span class="sxs-lookup"><span data-stu-id="00dde-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="00dde-138">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="00dde-138">**ItemClass**</span></span> | <span data-ttu-id="00dde-139">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-139">No</span></span> | <span data-ttu-id="00dde-p103">Указывает сопоставляемый специализированный класс сообщений. Дополнительные сведения см. в статье [Активация почтовой надстройки в Outlook для определенного класса сообщений](../../outlook/activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="00dde-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](../../outlook/activation-rules.md).</span></span> |
| <span data-ttu-id="00dde-142">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="00dde-142">**IncludeSubClasses**</span></span> | <span data-ttu-id="00dde-143">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-143">No</span></span> | <span data-ttu-id="00dde-144">Указывает, должно ли правило оцениваться как истинное (true), если элемент принадлежит к подклассу указанного класса сообщений; по умолчанию используется значение `false`.</span><span class="sxs-lookup"><span data-stu-id="00dde-144">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="00dde-145">Пример</span><span class="sxs-lookup"><span data-stu-id="00dde-145">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="00dde-146">Правило ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="00dde-146">ItemHasAttachment rule</span></span>

<span data-ttu-id="00dde-147">Определяет правило, которое оценивается как истинное, если элемент содержит вложение.</span><span class="sxs-lookup"><span data-stu-id="00dde-147">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="00dde-148">Пример</span><span class="sxs-lookup"><span data-stu-id="00dde-148">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="00dde-149">Правило ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="00dde-149">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="00dde-150">Определяет правило, которое оценивается как истинное, если элемент содержит текст указанного типа сущности в теме или основном тексте.</span><span class="sxs-lookup"><span data-stu-id="00dde-150">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="00dde-151">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00dde-151">Attributes</span></span>

| <span data-ttu-id="00dde-152">Атрибут</span><span class="sxs-lookup"><span data-stu-id="00dde-152">Attribute</span></span> | <span data-ttu-id="00dde-153">Обязательный</span><span class="sxs-lookup"><span data-stu-id="00dde-153">Required</span></span> | <span data-ttu-id="00dde-154">Описание</span><span class="sxs-lookup"><span data-stu-id="00dde-154">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="00dde-155">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="00dde-155">**EntityType**</span></span> | <span data-ttu-id="00dde-156">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-156">Yes</span></span> | <span data-ttu-id="00dde-p104">Указывает тип сущности, который должен обнаруживаться, чтобы правило было оценено как истинное. Допустимые значения: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` и `Contact`.</span><span class="sxs-lookup"><span data-stu-id="00dde-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="00dde-159">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="00dde-159">**RegExFilter**</span></span> | <span data-ttu-id="00dde-160">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-160">No</span></span> | <span data-ttu-id="00dde-161">Задает регулярное выражение, которое должно выполняться в этой сущности для активации.</span><span class="sxs-lookup"><span data-stu-id="00dde-161">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="00dde-162">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="00dde-162">**FilterName**</span></span> | <span data-ttu-id="00dde-163">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-163">No</span></span> | <span data-ttu-id="00dde-164">Задает имя фильтра регулярных выражений, чтобы на этот фильтр можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="00dde-164">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="00dde-165">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="00dde-165">**IgnoreCase**</span></span> | <span data-ttu-id="00dde-166">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-166">No</span></span> | <span data-ttu-id="00dde-167">Указывает, следует ли игнорировать регистр при сравнении регулярного выражения, заданного атрибутом **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="00dde-167">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="00dde-168">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="00dde-168">**Highlight**</span></span> | <span data-ttu-id="00dde-169">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-169">No</span></span> | <span data-ttu-id="00dde-p105">**Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующие сущности. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="00dde-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="00dde-174">Пример</span><span class="sxs-lookup"><span data-stu-id="00dde-174">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="00dde-175">Правило ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="00dde-175">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="00dde-176">Задает правило, которое оценивается как истинное, если в указанном свойстве элемента обнаруживается соответствие для указанного регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="00dde-176">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="00dde-177">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00dde-177">Attributes</span></span>

| <span data-ttu-id="00dde-178">Атрибут</span><span class="sxs-lookup"><span data-stu-id="00dde-178">Attribute</span></span> | <span data-ttu-id="00dde-179">Обязательный</span><span class="sxs-lookup"><span data-stu-id="00dde-179">Required</span></span> | <span data-ttu-id="00dde-180">Описание</span><span class="sxs-lookup"><span data-stu-id="00dde-180">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="00dde-181">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="00dde-181">**RegExName**</span></span> | <span data-ttu-id="00dde-182">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-182">Yes</span></span> | <span data-ttu-id="00dde-183">Указывает имя регулярного выражения, чтобы на него можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="00dde-183">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="00dde-184">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="00dde-184">**RegExValue**</span></span> | <span data-ttu-id="00dde-185">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-185">Yes</span></span> | <span data-ttu-id="00dde-186">Указывает регулярное выражение, которое будет вычислено, чтобы определить, требуется ли отображать надстройку.</span><span class="sxs-lookup"><span data-stu-id="00dde-186">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="00dde-187">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="00dde-187">**PropertyName**</span></span> | <span data-ttu-id="00dde-188">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-188">Yes</span></span> | <span data-ttu-id="00dde-p106">Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` и `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="00dde-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="00dde-191">Если вы укажете `BodyAsHTML`, Outlook будет применять регулярное выражение, только если текст элемента представлен в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="00dde-191">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="00dde-192">В противном случае Outlook возвращает отсутствие совпадений для этого регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="00dde-192">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="00dde-193">Если вы укажете `BodyAsPlaintext`, Outlook всегда будет применять регулярное выражение для текста элемента.</span><span class="sxs-lookup"><span data-stu-id="00dde-193">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="00dde-194">**Примечание.** Необходимо задать атрибут **PropertyName** для `BodyAsPlaintext`, если указан атрибут **Highlight** для элемента **Rule**.</span><span class="sxs-lookup"><span data-stu-id="00dde-194">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="00dde-195">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="00dde-195">**IgnoreCase**</span></span> | <span data-ttu-id="00dde-196">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-196">No</span></span> | <span data-ttu-id="00dde-197">Указывает, следует ли игнорировать регистр при сравнении регулярного выражения, заданного атрибутом **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="00dde-197">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="00dde-198">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="00dde-198">**Highlight**</span></span> | <span data-ttu-id="00dde-199">Нет</span><span class="sxs-lookup"><span data-stu-id="00dde-199">No</span></span> | <span data-ttu-id="00dde-200">Указывает, как клиент должен выделять соответствующий текст.</span><span class="sxs-lookup"><span data-stu-id="00dde-200">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="00dde-201">Этот атрибут может применяться только к элементам **Rule**, вложенным в элементы **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="00dde-201">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="00dde-202">Допустимые значения: `all` и `none`.</span><span class="sxs-lookup"><span data-stu-id="00dde-202">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="00dde-203">Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="00dde-203">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="00dde-204">**Примечание.** Необходимо задать атрибут **PropertyName** для `BodyAsPlaintext`, если указан атрибут **Highlight** для элемента **Rule**.</span><span class="sxs-lookup"><span data-stu-id="00dde-204">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="00dde-205">Пример</span><span class="sxs-lookup"><span data-stu-id="00dde-205">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="00dde-206">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="00dde-206">RuleCollection</span></span>

<span data-ttu-id="00dde-207">Задает коллекцию правил и логический оператор, который должен использоваться при их оценке.</span><span class="sxs-lookup"><span data-stu-id="00dde-207">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="00dde-208">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00dde-208">Attributes</span></span>

| <span data-ttu-id="00dde-209">Атрибут</span><span class="sxs-lookup"><span data-stu-id="00dde-209">Attribute</span></span> | <span data-ttu-id="00dde-210">Обязательный</span><span class="sxs-lookup"><span data-stu-id="00dde-210">Required</span></span> | <span data-ttu-id="00dde-211">Описание</span><span class="sxs-lookup"><span data-stu-id="00dde-211">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="00dde-212">**Mode**</span><span class="sxs-lookup"><span data-stu-id="00dde-212">**Mode**</span></span> | <span data-ttu-id="00dde-213">Да</span><span class="sxs-lookup"><span data-stu-id="00dde-213">Yes</span></span> | <span data-ttu-id="00dde-p109">Указывает логический оператор, используемый при оценке коллекции правил. Допустимые значения: `And` и `Or`.</span><span class="sxs-lookup"><span data-stu-id="00dde-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="00dde-216">Пример</span><span class="sxs-lookup"><span data-stu-id="00dde-216">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="00dde-217">См. также</span><span class="sxs-lookup"><span data-stu-id="00dde-217">See also</span></span>

- [<span data-ttu-id="00dde-218">Правила активации для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="00dde-218">Activation rules for Outlook add-ins</span></span>](../../outlook/activation-rules.md)
- [<span data-ttu-id="00dde-219">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="00dde-219">Match strings in an Outlook item as well-known entities</span></span>](../../outlook/match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="00dde-220">Использование регулярных правил активации выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="00dde-220">Use regular expression activation rules to show an Outlook add-in</span></span>](../../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
