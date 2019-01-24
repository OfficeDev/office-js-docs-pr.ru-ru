---
title: Элемент Rule в файле манифеста
description: ''
ms.date: 12/27/2018
localization_priority: Normal
ms.openlocfilehash: 38e724e6962c48efd0902be315c49ebb4cf6c798
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388502"
---
# <a name="rule-element"></a><span data-ttu-id="c78b4-102">Элемент Rule</span><span class="sxs-lookup"><span data-stu-id="c78b4-102">Rule element</span></span>

<span data-ttu-id="c78b4-103">Указывает правила активации, которые следует оценивать для этой контекстной почтовой надстройки.</span><span class="sxs-lookup"><span data-stu-id="c78b4-103">Specifies the activation rule(s) that should be evaluated for this contextual mail add-in.</span></span>

<span data-ttu-id="c78b4-104">**Тип надстройки:** контекстная почтовая надстройка</span><span class="sxs-lookup"><span data-stu-id="c78b4-104">**Add-in type:** Mail contextual add-in</span></span>

## <a name="contained-in"></a><span data-ttu-id="c78b4-105">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c78b4-105">Contained in</span></span>

- [<span data-ttu-id="c78b4-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c78b4-106">OfficeApp</span></span>](officeapp.md)
- [<span data-ttu-id="c78b4-107">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c78b4-107">ExtensionPoint</span></span>](extensionpoint.md)

## <a name="attributes"></a><span data-ttu-id="c78b4-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c78b4-108">Attributes</span></span>

| <span data-ttu-id="c78b4-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c78b4-109">Attribute</span></span> | <span data-ttu-id="c78b4-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c78b4-110">Required</span></span> | <span data-ttu-id="c78b4-111">Описание</span><span class="sxs-lookup"><span data-stu-id="c78b4-111">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c78b4-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c78b4-112">**xsi:type**</span></span> | <span data-ttu-id="c78b4-113">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-113">Yes</span></span> | <span data-ttu-id="c78b4-114">Тип определяемого правила.</span><span class="sxs-lookup"><span data-stu-id="c78b4-114">The type of rule being defined.</span></span> |

<span data-ttu-id="c78b4-115">Правило может относиться к одному из указанных ниже типов.</span><span class="sxs-lookup"><span data-stu-id="c78b4-115">The type of rule can be one of the following.</span></span>

- [<span data-ttu-id="c78b4-116">ItemIs</span><span class="sxs-lookup"><span data-stu-id="c78b4-116">ItemIs</span></span>](#itemis-rule)
- [<span data-ttu-id="c78b4-117">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="c78b4-117">ItemHasAttachment</span></span>](#itemhasattachment-rule)
- [<span data-ttu-id="c78b4-118">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="c78b4-118">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)
- [<span data-ttu-id="c78b4-119">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="c78b4-119">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)
- [<span data-ttu-id="c78b4-120">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="c78b4-120">RuleCollection</span></span>](#rulecollection)

## <a name="itemis-rule"></a><span data-ttu-id="c78b4-121">Правило ItemIs</span><span class="sxs-lookup"><span data-stu-id="c78b4-121">ItemIs rule</span></span>

<span data-ttu-id="c78b4-122">Определяет правило, которое оценивается как истинное, если выбранный элемент относится к указанному типу.</span><span class="sxs-lookup"><span data-stu-id="c78b4-122">Defines a rule that evaluates to true if the selected item is of the specified type.</span></span>

### <a name="attributes"></a><span data-ttu-id="c78b4-123">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c78b4-123">Attributes</span></span>

| <span data-ttu-id="c78b4-124">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c78b4-124">Attribute</span></span> | <span data-ttu-id="c78b4-125">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c78b4-125">Required</span></span> | <span data-ttu-id="c78b4-126">Описание</span><span class="sxs-lookup"><span data-stu-id="c78b4-126">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c78b4-127">**ItemType**</span><span class="sxs-lookup"><span data-stu-id="c78b4-127">**ItemType**</span></span> | <span data-ttu-id="c78b4-128">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-128">Yes</span></span> | <span data-ttu-id="c78b4-p101">Указывает сопоставляемый тип элемента. Допустимые значения: `Message` и `Appointment`. К типу элементов `Message` относятся электронные письма, приглашения на собрания, ответы на них и уведомления об их отмене.</span><span class="sxs-lookup"><span data-stu-id="c78b4-p101">Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations.</span></span> |
| <span data-ttu-id="c78b4-132">**FormType**</span><span class="sxs-lookup"><span data-stu-id="c78b4-132">**FormType**</span></span> | <span data-ttu-id="c78b4-133">Нет (в [ExtensionPoint](extensionpoint.md)), да (в [OfficeApp](officeapp.md))</span><span class="sxs-lookup"><span data-stu-id="c78b4-133">No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md))</span></span> | <span data-ttu-id="c78b4-p102">Указывает, должно ли приложение отображаться в форме чтения или редактирования элемента. Допустимые значения: `Read`, `Edit`, `ReadOrEdit`. Для объекта `Rule` в `ExtensionPoint` НЕОБХОДИМО использовать значение `Read`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-p102">Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`.</span></span> |
| <span data-ttu-id="c78b4-137">**ItemClass**</span><span class="sxs-lookup"><span data-stu-id="c78b4-137">**ItemClass**</span></span> | <span data-ttu-id="c78b4-138">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-138">No</span></span> | <span data-ttu-id="c78b4-p103">Указывает сопоставляемый специализированный класс сообщений. Дополнительные сведения см. в статье [Активация почтовой надстройки в Outlook для определенного класса сообщений](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="c78b4-p103">Specifies the custom message class to match. For more information, see [Activate a mail add-in in Outlook for a specific message class](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> |
| <span data-ttu-id="c78b4-141">**IncludeSubClasses**</span><span class="sxs-lookup"><span data-stu-id="c78b4-141">**IncludeSubClasses**</span></span> | <span data-ttu-id="c78b4-142">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-142">No</span></span> | <span data-ttu-id="c78b4-143">Указывает, должно ли правило оцениваться как истинное (true), если элемент принадлежит к подклассу указанного класса сообщений; по умолчанию используется значение `false`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-143">Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`.</span></span> |

### <a name="example"></a><span data-ttu-id="c78b4-144">Пример</span><span class="sxs-lookup"><span data-stu-id="c78b4-144">Example</span></span>

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## <a name="itemhasattachment-rule"></a><span data-ttu-id="c78b4-145">Правило ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="c78b4-145">ItemHasAttachment rule</span></span>

<span data-ttu-id="c78b4-146">Определяет правило, которое оценивается как истинное, если элемент содержит вложение.</span><span class="sxs-lookup"><span data-stu-id="c78b4-146">Defines a rule that evaluates to true if the item contains an attachment.</span></span>

### <a name="example"></a><span data-ttu-id="c78b4-147">Пример</span><span class="sxs-lookup"><span data-stu-id="c78b4-147">Example</span></span>

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="c78b4-148">Правило ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="c78b4-148">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="c78b4-149">Определяет правило, которое оценивается как истинное, если элемент содержит текст указанного типа сущности в теме или основном тексте.</span><span class="sxs-lookup"><span data-stu-id="c78b4-149">Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.</span></span>

### <a name="attributes"></a><span data-ttu-id="c78b4-150">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c78b4-150">Attributes</span></span>

| <span data-ttu-id="c78b4-151">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c78b4-151">Attribute</span></span> | <span data-ttu-id="c78b4-152">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c78b4-152">Required</span></span> | <span data-ttu-id="c78b4-153">Описание</span><span class="sxs-lookup"><span data-stu-id="c78b4-153">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c78b4-154">**EntityType**</span><span class="sxs-lookup"><span data-stu-id="c78b4-154">**EntityType**</span></span> | <span data-ttu-id="c78b4-155">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-155">Yes</span></span> | <span data-ttu-id="c78b4-p104">Указывает тип сущности, который должен обнаруживаться, чтобы правило было оценено как истинное. Допустимые значения: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress` и `Contact`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-p104">Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`.</span></span> |
| <span data-ttu-id="c78b4-158">**RegExFilter**</span><span class="sxs-lookup"><span data-stu-id="c78b4-158">**RegExFilter**</span></span> | <span data-ttu-id="c78b4-159">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-159">No</span></span> | <span data-ttu-id="c78b4-160">Задает регулярное выражение, которое должно выполняться в этой сущности для активации.</span><span class="sxs-lookup"><span data-stu-id="c78b4-160">Specifies a regular expression to run against this entity for activation.</span></span> |
| <span data-ttu-id="c78b4-161">**FilterName**</span><span class="sxs-lookup"><span data-stu-id="c78b4-161">**FilterName**</span></span> | <span data-ttu-id="c78b4-162">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-162">No</span></span> | <span data-ttu-id="c78b4-163">Задает имя фильтра регулярных выражений, чтобы на этот фильтр можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="c78b4-163">Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code.</span></span> |
| <span data-ttu-id="c78b4-164">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="c78b4-164">**IgnoreCase**</span></span> | <span data-ttu-id="c78b4-165">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-165">No</span></span> | <span data-ttu-id="c78b4-166">Указывает, следует ли игнорировать регистр при сравнении регулярного выражения, заданного атрибутом **RegExFilter**.</span><span class="sxs-lookup"><span data-stu-id="c78b4-166">Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute.</span></span> |
| <span data-ttu-id="c78b4-167">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="c78b4-167">**Highlight**</span></span> | <span data-ttu-id="c78b4-168">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-168">No</span></span> | <span data-ttu-id="c78b4-p105">**Примечание.** Это относится только к элементам **Rule**, вложенным в элементы **ExtensionPoint**. Указывает, как клиент должен выделять соответствующие сущности. Допустимые значения: `all` и `none`. Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-p105">**Note:** this only applies to **Rule** elements within **ExtensionPoint** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.</span></span> |

### <a name="example"></a><span data-ttu-id="c78b4-173">Пример</span><span class="sxs-lookup"><span data-stu-id="c78b4-173">Example</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="c78b4-174">Правило ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="c78b4-174">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="c78b4-175">Задает правило, которое оценивается как истинное, если в указанном свойстве элемента обнаруживается соответствие для указанного регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="c78b4-175">Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item.</span></span>

### <a name="attributes"></a><span data-ttu-id="c78b4-176">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c78b4-176">Attributes</span></span>

| <span data-ttu-id="c78b4-177">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c78b4-177">Attribute</span></span> | <span data-ttu-id="c78b4-178">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c78b4-178">Required</span></span> | <span data-ttu-id="c78b4-179">Описание</span><span class="sxs-lookup"><span data-stu-id="c78b4-179">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c78b4-180">**RegExName**</span><span class="sxs-lookup"><span data-stu-id="c78b4-180">**RegExName**</span></span> | <span data-ttu-id="c78b4-181">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-181">Yes</span></span> | <span data-ttu-id="c78b4-182">Указывает имя регулярного выражения, чтобы на него можно было ссылаться в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="c78b4-182">Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in.</span></span> |
| <span data-ttu-id="c78b4-183">**RegExValue**</span><span class="sxs-lookup"><span data-stu-id="c78b4-183">**RegExValue**</span></span> | <span data-ttu-id="c78b4-184">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-184">Yes</span></span> | <span data-ttu-id="c78b4-185">Указывает регулярное выражение, которое будет вычислено, чтобы определить, требуется ли отображать надстройку.</span><span class="sxs-lookup"><span data-stu-id="c78b4-185">Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown.</span></span> |
| <span data-ttu-id="c78b4-186">**PropertyName**</span><span class="sxs-lookup"><span data-stu-id="c78b4-186">**PropertyName**</span></span> | <span data-ttu-id="c78b4-187">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-187">Yes</span></span> | <span data-ttu-id="c78b4-p106">Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения. Допустимые значения: `Subject`, `BodyAsPlaintext`, `BodyAsHTML` и `SenderSMTPAddress`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-p106">Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.</span></span><br/><br/><span data-ttu-id="c78b4-190">Если вы укажете `BodyAsHTML`, Outlook будет применять регулярное выражение, только если текст элемента представлен в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c78b4-190">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="c78b4-191">В противном случае Outlook возвращает отсутствие совпадений для этого регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="c78b4-191">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="c78b4-192">Если вы укажете `BodyAsPlaintext`, Outlook всегда будет применять регулярное выражение для текста элемента.</span><span class="sxs-lookup"><span data-stu-id="c78b4-192">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="c78b4-193">**Примечание.** Необходимо задать атрибут **PropertyName** для `BodyAsPlaintext`, если указан атрибут **Highlight** для элемента **Rule**.</span><span class="sxs-lookup"><span data-stu-id="c78b4-193">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>|
| <span data-ttu-id="c78b4-194">**IgnoreCase**</span><span class="sxs-lookup"><span data-stu-id="c78b4-194">**IgnoreCase**</span></span> | <span data-ttu-id="c78b4-195">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-195">No</span></span> | <span data-ttu-id="c78b4-196">Указывает, следует ли игнорировать регистр при сравнении регулярного выражения, заданного атрибутом **RegExName**.</span><span class="sxs-lookup"><span data-stu-id="c78b4-196">Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute.</span></span> |
| <span data-ttu-id="c78b4-197">**Highlight**</span><span class="sxs-lookup"><span data-stu-id="c78b4-197">**Highlight**</span></span> | <span data-ttu-id="c78b4-198">Нет</span><span class="sxs-lookup"><span data-stu-id="c78b4-198">No</span></span> | <span data-ttu-id="c78b4-199">Указывает, как клиент должен выделять соответствующий текст.</span><span class="sxs-lookup"><span data-stu-id="c78b4-199">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="c78b4-200">Этот атрибут может применяться только к элементам **Rule**, вложенным в элементы **ExtensionPoint**.</span><span class="sxs-lookup"><span data-stu-id="c78b4-200">This attribute can only be applied to **Rule** elements within **ExtensionPoint** elements.</span></span> <span data-ttu-id="c78b4-201">Допустимые значения: `all` и `none`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-201">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="c78b4-202">Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-202">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="c78b4-203">**Примечание.** Необходимо задать атрибут **PropertyName** для `BodyAsPlaintext`, если указан атрибут **Highlight** для элемента **Rule**.</span><span class="sxs-lookup"><span data-stu-id="c78b4-203">**Note:** You must set the **PropertyName** attribute to `BodyAsPlaintext` if you specify the **Highlight** attribute for the **Rule** element.</span></span>
|

### <a name="example"></a><span data-ttu-id="c78b4-204">Пример</span><span class="sxs-lookup"><span data-stu-id="c78b4-204">Example</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsHTML" IgnoreCase="true" />
```

## <a name="rulecollection"></a><span data-ttu-id="c78b4-205">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="c78b4-205">RuleCollection</span></span>

<span data-ttu-id="c78b4-206">Задает коллекцию правил и логический оператор, который должен использоваться при их оценке.</span><span class="sxs-lookup"><span data-stu-id="c78b4-206">Defines a collection of rules and the logical operator to use when evaluating them.</span></span>

### <a name="attributes"></a><span data-ttu-id="c78b4-207">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c78b4-207">Attributes</span></span>

| <span data-ttu-id="c78b4-208">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c78b4-208">Attribute</span></span> | <span data-ttu-id="c78b4-209">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c78b4-209">Required</span></span> | <span data-ttu-id="c78b4-210">Описание</span><span class="sxs-lookup"><span data-stu-id="c78b4-210">Description</span></span> |
|:-----|:-----|:-----|
| <span data-ttu-id="c78b4-211">**Mode**</span><span class="sxs-lookup"><span data-stu-id="c78b4-211">**Mode**</span></span> | <span data-ttu-id="c78b4-212">Да</span><span class="sxs-lookup"><span data-stu-id="c78b4-212">Yes</span></span> | <span data-ttu-id="c78b4-p109">Указывает логический оператор, используемый при оценке коллекции правил. Допустимые значения: `And` и `Or`.</span><span class="sxs-lookup"><span data-stu-id="c78b4-p109">Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`.</span></span> |

### <a name="example"></a><span data-ttu-id="c78b4-215">Пример</span><span class="sxs-lookup"><span data-stu-id="c78b4-215">Example</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" ItemType="Message" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
</Rule>
```

## <a name="see-also"></a><span data-ttu-id="c78b4-216">См. также</span><span class="sxs-lookup"><span data-stu-id="c78b4-216">See also</span></span>

- [<span data-ttu-id="c78b4-217">Правила активации для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="c78b4-217">Activation rules for Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/activation-rules)
- [<span data-ttu-id="c78b4-218">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="c78b4-218">Match strings in an Outlook item as well-known entities</span></span>](https://docs.microsoft.com/outlook/add-ins/match-strings-in-an-item-as-well-known-entities)    
- [<span data-ttu-id="c78b4-219">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="c78b4-219">Use regular expression activation rules to show an Outlook add-in</span></span>](https://docs.microsoft.com/outlook/add-ins/use-regular-expressions-to-show-an-outlook-add-in)
