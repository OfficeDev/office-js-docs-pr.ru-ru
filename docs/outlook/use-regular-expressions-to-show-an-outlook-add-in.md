---
title: Использование правил активации на основе регулярных выражений для отображения надстройки
description: Узнайте, как использовать правила активации на основе регулярных выражений для контекстных надстроек Outlook.
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: 4a5507b410ed729f76c3efa0119e87c6a6dbc71a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292477"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a><span data-ttu-id="0f909-103">Использование правил активации на основе регулярных выражений для отображения надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="0f909-103">Use regular expression activation rules to show an Outlook add-in</span></span>

<span data-ttu-id="0f909-104">Вы можете указать правила на основе регулярных выражений для активации [контекстной надстройки](contextual-outlook-add-ins.md) при обнаружении соответствия в определенных полях сообщения.</span><span class="sxs-lookup"><span data-stu-id="0f909-104">You can specify regular expression rules to have a [contextual add-in](contextual-outlook-add-ins.md) activated when a match is found in specific fields of the message.</span></span> <span data-ttu-id="0f909-105">Контекстные надстройки активируются только в режиме чтения. Outlook не активирует контекстные надстройки, когда пользователь создает элемент.</span><span class="sxs-lookup"><span data-stu-id="0f909-105">Contextual add-ins activate only in read mode, Outlook does not activate contextual add-ins when the user is composing an item.</span></span> <span data-ttu-id="0f909-106">Кроме того, существуют другие сценарии, в которых Outlook не активирует надстройки, например элементы с цифровой подписью.</span><span class="sxs-lookup"><span data-stu-id="0f909-106">There are also other scenarios where Outlook does not activate add-ins, for example, digitally signed items.</span></span> <span data-ttu-id="0f909-107">Дополнительные сведения см. в статье [Правила активации для надстроек Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="0f909-107">For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="0f909-108">Вы можете указать регулярное выражение в составе правила [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) или [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) в XML-файле манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="0f909-108">You can specify a regular expression as part of an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule or [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule in the add-in XML manifest.</span></span> <span data-ttu-id="0f909-109">Правила указываются в точке расширения [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity).</span><span class="sxs-lookup"><span data-stu-id="0f909-109">The rules are specified in a [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) extension point.</span></span>

<span data-ttu-id="0f909-110">Outlook оценивает регулярные выражения на основе правил для интерпретатора JavaScript, используемых браузером на клиентском компьютере.</span><span class="sxs-lookup"><span data-stu-id="0f909-110">Outlook evaluates regular expressions based on the rules for the JavaScript interpreter used by the browser on the client computer.</span></span> <span data-ttu-id="0f909-111">Outlook поддерживает те же специальные знаки, что и все обработчики XML.</span><span class="sxs-lookup"><span data-stu-id="0f909-111">Outlook supports the same list of special characters that all XML processors also support.</span></span> <span data-ttu-id="0f909-112">Они перечислены в следующей таблице.</span><span class="sxs-lookup"><span data-stu-id="0f909-112">The following table lists these special characters.</span></span> <span data-ttu-id="0f909-113">Указывая эти знаки в регулярных выражениях, используйте соответствующие escape-последовательности из следующей таблицы.</span><span class="sxs-lookup"><span data-stu-id="0f909-113">You can use these characters in a regular expression by specifying the escaped sequence for the corresponding character, as described in the following table.</span></span>

<br/>

|<span data-ttu-id="0f909-114">Знак</span><span class="sxs-lookup"><span data-stu-id="0f909-114">Character</span></span>|<span data-ttu-id="0f909-115">Описание</span><span class="sxs-lookup"><span data-stu-id="0f909-115">Description</span></span>|<span data-ttu-id="0f909-116">Escape-последовательность</span><span class="sxs-lookup"><span data-stu-id="0f909-116">Escape sequence to use</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="0f909-117">Двойная кавычка</span><span class="sxs-lookup"><span data-stu-id="0f909-117">Double quotation mark</span></span>|`&quot;`|
|`&`|<span data-ttu-id="0f909-118">Амперсанд</span><span class="sxs-lookup"><span data-stu-id="0f909-118">Ampersand</span></span>|`&amp;`|
|`'`|<span data-ttu-id="0f909-119">Апостроф</span><span class="sxs-lookup"><span data-stu-id="0f909-119">Apostrophe</span></span>|`&apos;`|
|`<`|<span data-ttu-id="0f909-120">Знак "меньше"</span><span class="sxs-lookup"><span data-stu-id="0f909-120">Less-than sign</span></span>|`&lt;`|
|`>`|<span data-ttu-id="0f909-121">Знак "больше"</span><span class="sxs-lookup"><span data-stu-id="0f909-121">Greater-than sign</span></span>|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="0f909-122">Правило ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="0f909-122">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="0f909-123">Правило `ItemHasRegularExpressionMatch` позволяет управлять активацией надстройки в зависимости от определенных значений поддерживаемого свойства.</span><span class="sxs-lookup"><span data-stu-id="0f909-123">An  `ItemHasRegularExpressionMatch` rule is useful in controlling activation of an add-in based on specific values of a supported property.</span></span> <span data-ttu-id="0f909-124">Ниже описаны атрибуты правила `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="0f909-124">The `ItemHasRegularExpressionMatch` rule has the following attributes.</span></span>

<br/>

|<span data-ttu-id="0f909-125">Имя атрибута</span><span class="sxs-lookup"><span data-stu-id="0f909-125">Attribute name</span></span>|<span data-ttu-id="0f909-126">Описание</span><span class="sxs-lookup"><span data-stu-id="0f909-126">Description</span></span>|
|:-----|:-----|
|`RegExName`|<span data-ttu-id="0f909-127">Указывает имя регулярного выражения, чтобы вы могли сослаться на него в коде надстройки.</span><span class="sxs-lookup"><span data-stu-id="0f909-127">Specifies the name of the regular expression so that you can refer to the expression in the code for your add-in.</span></span>|
|`RegExValue`|<span data-ttu-id="0f909-128">Указывает регулярное выражение, которое будет рассчитано для определения необходимости отображения надстройки.</span><span class="sxs-lookup"><span data-stu-id="0f909-128">Specifies the regular expression that will be evaluated to determine whether the add-in should be shown.</span></span>|
|`PropertyName`|<span data-ttu-id="0f909-129">Указывает имя свойства, которое будет использоваться для вычисления регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="0f909-129">Specifies the name of the property that the regular expression will be evaluated against.</span></span> <span data-ttu-id="0f909-130">Допустимые значения — `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress` и `Subject`.</span><span class="sxs-lookup"><span data-stu-id="0f909-130">The allowed values are `BodyAsHTML`, `BodyAsPlaintext`, `SenderSMTPAddress`, and `Subject`.</span></span><br/><br/><span data-ttu-id="0f909-131">Если вы укажете `BodyAsHTML`, Outlook будет применять регулярное выражение, только если текст элемента представлен в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="0f909-131">If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML.</span></span> <span data-ttu-id="0f909-132">В противном случае Outlook возвращает отсутствие совпадений для этого регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="0f909-132">Otherwise, Outlook returns no matches for that regular expression.</span></span><br/><br/><span data-ttu-id="0f909-133">Если вы укажете `BodyAsPlaintext`, Outlook всегда будет применять регулярное выражение для текста элемента.</span><span class="sxs-lookup"><span data-stu-id="0f909-133">If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.</span></span><br/><br/><span data-ttu-id="0f909-134">**Примечание.** Необходимо задать атрибут `PropertyName` для `BodyAsPlaintext`, если указан атрибут `Highlight` для элемента `Rule`.</span><span class="sxs-lookup"><span data-stu-id="0f909-134">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span>|
|`IgnoreCase`|<span data-ttu-id="0f909-135">Указывает, следует ли игнорировать регистр при поиске соответствий регулярному выражению, заданному атрибутом `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="0f909-135">Specifies whether to ignore case when matching the regular expression specified by `RegExName`.</span></span>|
| `Highlight` | <span data-ttu-id="0f909-136">Указывает, как клиент должен выделять соответствующий текст.</span><span class="sxs-lookup"><span data-stu-id="0f909-136">Specifies how the client should highlight matching text.</span></span> <span data-ttu-id="0f909-137">Этот элемент может применяться только к элементам `Rule`, вложенным в элементы `ExtensionPoint`.</span><span class="sxs-lookup"><span data-stu-id="0f909-137">This element can only be applied to `Rule` elements within `ExtensionPoint` elements.</span></span> <span data-ttu-id="0f909-138">Допустимые значения: `all` и `none`.</span><span class="sxs-lookup"><span data-stu-id="0f909-138">Can be one of the following: `all` or `none`.</span></span> <span data-ttu-id="0f909-139">Если этот атрибут не задан, по умолчанию используется значение `all`.</span><span class="sxs-lookup"><span data-stu-id="0f909-139">If not specified, the default value is `all`.</span></span><br/><br/><span data-ttu-id="0f909-140">**Примечание.** Необходимо задать атрибут `PropertyName` для `BodyAsPlaintext`, если указан атрибут `Highlight` для элемента `Rule`.</span><span class="sxs-lookup"><span data-stu-id="0f909-140">**Note:** You must set the `PropertyName` attribute to `BodyAsPlaintext` if you specify the `Highlight` attribute for the `Rule` element.</span></span> |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a><span data-ttu-id="0f909-141">Рекомендации по использованию регулярных выражений в правилах</span><span class="sxs-lookup"><span data-stu-id="0f909-141">Best practices for using regular expressions in rules</span></span>

<span data-ttu-id="0f909-142">При использовании регулярных выражений уделяйте особое внимание следующим аспектам:</span><span class="sxs-lookup"><span data-stu-id="0f909-142">Pay special attention to the following when you use regular expressions:</span></span>

- <span data-ttu-id="0f909-143">Если вы указываете правило `ItemHasRegularExpressionMatch` для текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента.</span><span class="sxs-lookup"><span data-stu-id="0f909-143">If you specify an `ItemHasRegularExpressionMatch` rule on the body of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item.</span></span> <span data-ttu-id="0f909-144">Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="0f909-144">Using a regular expression such as `.*` to attempt to obtain the entire body of an item does not always return the expected results.</span></span>
- <span data-ttu-id="0f909-145">Возвращаемый обычный текст может несколько отличаться в зависимости браузера.</span><span class="sxs-lookup"><span data-stu-id="0f909-145">The plain text body returned on one browser can be different in subtle ways on another.</span></span> <span data-ttu-id="0f909-146">Если вы используете правило `ItemHasRegularExpressionMatch` с таким значением атрибута `PropertyName`: `BodyAsPlaintext`, проверьте свое регулярное выражение во всех поддерживаемых надстройкой браузерах.</span><span class="sxs-lookup"><span data-stu-id="0f909-146">If you use an `ItemHasRegularExpressionMatch` rule with `BodyAsPlaintext` as the `PropertyName` attribute, test your regular expression on all the browsers that your add-in supports.</span></span>

    <span data-ttu-id="0f909-147">Так как в разных браузерах основной текст выбранного элемента считывается разными способами, ваше регулярное выражение должно учитывать мелкие различия, которые могут быть возвращены в составе основного текста.</span><span class="sxs-lookup"><span data-stu-id="0f909-147">Because different browsers use different ways to obtain the text body of a selected item, you should make sure that your regular expression supports the subtle differences that can be returned as part of the body text.</span></span> <span data-ttu-id="0f909-148">Например, в некоторых браузерах, таких как Internet Explorer 9, для получения основного текста элемента используется свойство `innerText` модели DOM, а в других (например, Firefox) — метод `.textContent()`.</span><span class="sxs-lookup"><span data-stu-id="0f909-148">For example, some browsers such as Internet Explorer 9 uses the `innerText` property of the DOM, and others such as Firefox uses the `.textContent()` method to obtain the text body of an item.</span></span> <span data-ttu-id="0f909-149">Кроме того, различные браузеры могут по-разному возвращать разрывы строк (в Internet Explorer — `\r\n`, а в Firefox и Chrome — `\n`).</span><span class="sxs-lookup"><span data-stu-id="0f909-149">Also, different browsers may return line breaks differently: a line break is `\r\n` on Internet Explorer, and `\n` on Firefox and Chrome.</span></span> <span data-ttu-id="0f909-150">Дополнительные сведения см. в документе [Консорциум W3C: совместимость с моделью DOM (HTML)](https://quirksmode.org/dom/html/).</span><span class="sxs-lookup"><span data-stu-id="0f909-150">For more information, se [W3C DOM Compatibility - HTML](https://quirksmode.org/dom/html/).</span></span>

- <span data-ttu-id="0f909-151">Текст элемента в HTML-формате немного отличается для полнофункционального клиента Outlook, Outlook в Интернете и Outlook для мобильных устройств.</span><span class="sxs-lookup"><span data-stu-id="0f909-151">The HTML body of an item is slightly different between an Outlook rich client, and Outlook on the web or Outlook mobile.</span></span> <span data-ttu-id="0f909-152">Будьте внимательны, задавая регулярные выражения.</span><span class="sxs-lookup"><span data-stu-id="0f909-152">Define your regular expressions carefully.</span></span>

- <span data-ttu-id="0f909-p112">В зависимости от клиента Outlook, типа устройства или свойства, к которому применяется регулярное выражение, существуют другие рекомендации и пределы для каждого из клиентов, которые следует учитывать при разработке регулярных выражений в качестве правил активации. Для получения дополнительных сведений ознакомьтесь с разделом об [ограничении для активации и API JavaScript для надстроек Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) .</span><span class="sxs-lookup"><span data-stu-id="0f909-p112">Depending on the Outlook client, type of device, or property that a regular expression is being applied on, there are other best practices and limits for each of the clients that you should be aware of when designing regular expressions as activation rules. See [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) for details.</span></span>

### <a name="examples"></a><span data-ttu-id="0f909-155">Примеры</span><span class="sxs-lookup"><span data-stu-id="0f909-155">Examples</span></span>

<span data-ttu-id="0f909-156">Следующее правило `ItemHasRegularExpressionMatch` активирует надстройку, если SMTP-адрес отправителя содержит строку `@contoso` без учета регистра.</span><span class="sxs-lookup"><span data-stu-id="0f909-156">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever the sender's SMTP email address matches `@contoso`, regardless of uppercase or lowercase characters.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

<span data-ttu-id="0f909-157">Ниже приведен другой способ указания того же регулярного выражения с использованием атрибута `IgnoreCase`.</span><span class="sxs-lookup"><span data-stu-id="0f909-157">The following is another way to specify the same regular expression using the  `IgnoreCase` attribute.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

<span data-ttu-id="0f909-158">Следующее правило `ItemHasRegularExpressionMatch` активирует надстройку, если основной текст текущего элемента содержит биржевой символ акции.</span><span class="sxs-lookup"><span data-stu-id="0f909-158">The following `ItemHasRegularExpressionMatch` rule activates the add-in whenever a stock symbol is included in the body of the current item.</span></span>

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a><span data-ttu-id="0f909-159">Правило ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="0f909-159">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="0f909-160">Правило `ItemHasKnownEntity` активирует надстройку при наличии сущности в теме или тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="0f909-160">An `ItemHasKnownEntity` rule activates an add-in based on the existence of an entity in the subject or body of the selected item.</span></span> <span data-ttu-id="0f909-161">Тип [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) определяет поддерживаемые сущности.</span><span class="sxs-lookup"><span data-stu-id="0f909-161">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) type defines the supported entities.</span></span> <span data-ttu-id="0f909-162">Применять регулярное выражение в правиле `ItemHasKnownEntity` удобно, когда активация надстройки зависит от группы значений сущности (например, определенного набора URL-адресов или номеров телефонов с определенным кодом области).</span><span class="sxs-lookup"><span data-stu-id="0f909-162">Applying a regular expression on an `ItemHasKnownEntity` rule provides the convenience where activation is based on a subset of values for an entity (for example, a specific set of URLs, or telephone numbers with a certain area code).</span></span>

> [!NOTE]
> <span data-ttu-id="0f909-163">Независимо от языкового стандарта, указанного в манифесте, Outlook может извлекать строки сущностей только на английском языке.</span><span class="sxs-lookup"><span data-stu-id="0f909-163">Outlook can only extract entity strings in English regardless of the default locale specified in the manifest.</span></span> <span data-ttu-id="0f909-164">Только сообщения поддерживают тип сущности `MeetingSuggestion`.</span><span class="sxs-lookup"><span data-stu-id="0f909-164">Only messages support the `MeetingSuggestion` entity type; appointments do not.</span></span> <span data-ttu-id="0f909-165">Сущности невозможно извлечь из элементов в папке **Отправленные**. Правило `ItemHasKnownEntity` не подходит для активации надстройки для элементов в папке **Отправленные**.</span><span class="sxs-lookup"><span data-stu-id="0f909-165">You cannot extract entities from items in the **Sent Items** folder, nor can you use an `ItemHasKnownEntity` rule to activate an add-in for items in the **Sent Items** folder.</span></span>

<span data-ttu-id="0f909-166">Правило `ItemHasKnownEntity` поддерживает атрибуты, перечисленные в следующей таблице.</span><span class="sxs-lookup"><span data-stu-id="0f909-166">The `ItemHasKnownEntity` rule supports the attributes in the following table.</span></span> <span data-ttu-id="0f909-167">Обратите внимание, что указывать регулярное выражение в правиле `ItemHasKnownEntity` необязательно, но при использовании регулярного выражения в качестве фильтра сущности необходимо указывать атрибуты `RegExFilter` и `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="0f909-167">Note that while specifying a regular expression is optional in an `ItemHasKnownEntity` rule, if you choose to use a regular expression as an entity filter, you must specify both the `RegExFilter` and `FilterName` attributes.</span></span>

<br/>

|<span data-ttu-id="0f909-168">Имя атрибута</span><span class="sxs-lookup"><span data-stu-id="0f909-168">Attribute name</span></span>|<span data-ttu-id="0f909-169">Описание</span><span class="sxs-lookup"><span data-stu-id="0f909-169">Description</span></span>|
|:-----|:-----|
|`EntityType`|<span data-ttu-id="0f909-170">Задает тип сущности, который должен быть обнаружен, чтобы правило было оценено как `true`.</span><span class="sxs-lookup"><span data-stu-id="0f909-170">Specifies the type of entity that must be found for the rule to evaluate to `true`.</span></span> <span data-ttu-id="0f909-171">Используйте несколько правил, чтобы указать несколько типов сущностей.</span><span class="sxs-lookup"><span data-stu-id="0f909-171">Use multiple rules to specify multiple types of entities.</span></span>|
|`RegExFilter`|<span data-ttu-id="0f909-172">Указывает регулярное выражение, обеспечивающее дальнейшую фильтрацию экземпляров сущности, указанной атрибутом `EntityType`.</span><span class="sxs-lookup"><span data-stu-id="0f909-172">Specifies a regular expression that further filters instances of the entity specified by `EntityType`.</span></span>|
|`FilterName`|<span data-ttu-id="0f909-173">Указывает имя регулярного выражения, заданного атрибутом `RegExFilter`, чтобы впоследствии можно было сослаться на него в коде.</span><span class="sxs-lookup"><span data-stu-id="0f909-173">Specifies the name of the regular expression specified by `RegExFilter`, so that it is subsequently possible to refer to it by code.</span></span>|
|`IgnoreCase`|<span data-ttu-id="0f909-174">Указывает, следует ли игнорировать регистр при поиске соответствий регулярному выражению, заданному атрибутом `RegExFilter`.</span><span class="sxs-lookup"><span data-stu-id="0f909-174">Specifies whether to ignore case when matching the regular expression specified by `RegExFilter`.</span></span>|

### <a name="examples"></a><span data-ttu-id="0f909-175">Примеры</span><span class="sxs-lookup"><span data-stu-id="0f909-175">Examples</span></span>

<span data-ttu-id="0f909-176">В следующем правиле `ItemHasKnownEntity` активация надстройки выполняется при наличии URL-адреса в теме или основном тексте текущего элемента и строки `youtube` в этом адресе независимо от регистра.</span><span class="sxs-lookup"><span data-stu-id="0f909-176">The following `ItemHasKnownEntity` rule activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string `youtube`, regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a><span data-ttu-id="0f909-177">Использование результатов регулярных выражений в коде</span><span class="sxs-lookup"><span data-stu-id="0f909-177">Using regular expression results in code</span></span>

<span data-ttu-id="0f909-178">Вы можете получить соответствия регулярному выражению, воспользовавшись следующими методами текущего элемента:</span><span class="sxs-lookup"><span data-stu-id="0f909-178">You can obtain matches to a regular expression by using the following methods on the current item:</span></span>

- <span data-ttu-id="0f909-179">Метод [getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) возвращает строки текущего элемента, соответствующие всем регулярным выражениям, указанным в правилах `ItemHasRegularExpressionMatch` и `ItemHasKnownEntity` для надстройки.</span><span class="sxs-lookup"><span data-stu-id="0f909-179">[getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for all regular expressions specified in `ItemHasRegularExpressionMatch` and `ItemHasKnownEntity` rules of the add-in.</span></span>

- <span data-ttu-id="0f909-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) возвращает строки текущего элемента, соответствующие определенному регулярному выражению, указанному в правиле `ItemHasRegularExpressionMatch` надстройки.</span><span class="sxs-lookup"><span data-stu-id="0f909-180">[getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns matches in the current item for the identified regular expression specified in an `ItemHasRegularExpressionMatch` rule of the add-in.</span></span>

- <span data-ttu-id="0f909-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) возвращает полные экземпляры сущностей, которые содержат соответствия определенному регулярному выражению, указанному в правиле `ItemHasKnownEntity` надстройки.</span><span class="sxs-lookup"><span data-stu-id="0f909-181">[getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) returns entire instances of entities that contain matches for the identified regular expression specified in an `ItemHasKnownEntity` rule of the add-in.</span></span>

<span data-ttu-id="0f909-182">При оценке регулярных выражений соответствия возвращаются в надстройку в массиве.</span><span class="sxs-lookup"><span data-stu-id="0f909-182">When the regular expressions are evaluated, the matches are returned to your add-in in an array object.</span></span> <span data-ttu-id="0f909-183">При использовании метода `getRegExMatches` идентификатор этого массива соответствует имени регулярного выражения.</span><span class="sxs-lookup"><span data-stu-id="0f909-183">For `getRegExMatches`, that object has the identifier of the name of the regular expression.</span></span>

> [!NOTE]
> <span data-ttu-id="0f909-184">Outlook не возвращает соответствия в каком-либо определенном порядке в массиве.</span><span class="sxs-lookup"><span data-stu-id="0f909-184">Outlook does not return matches in any particular order in the array.</span></span> <span data-ttu-id="0f909-185">Кроме того, соответствия могут возвращаться в другом порядке, даже если вы запустите ту же настройку в каждом из этих клиентов для того же элемента в том же почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="0f909-185">Also, you should not assume that matches are returned in the same order in this array even when you run the same add-in on each of these clients on the same item in the same mailbox.</span></span>

### <a name="examples"></a><span data-ttu-id="0f909-186">Примеры</span><span class="sxs-lookup"><span data-stu-id="0f909-186">Examples</span></span>

<span data-ttu-id="0f909-187">Ниже приведен пример коллекции правил, содержащей правило `ItemHasRegularExpressionMatch` с регулярным выражением `videoURL`.</span><span class="sxs-lookup"><span data-stu-id="0f909-187">The following is an example of a rule collection that contains an  `ItemHasRegularExpressionMatch` rule with a regular expression named `videoURL`.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

<span data-ttu-id="0f909-188">В следующем примере используется метод `getRegExMatches` текущего элемента, чтобы поместить в переменную `videos` результаты предыдущего правила `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="0f909-188">The following example uses `getRegExMatches` of the current item to set a variable `videos` to the results of the preceding `ItemHasRegularExpressionMatch` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

<span data-ttu-id="0f909-p119">Несколько совпадений хранятся в этом объекте в виде элементов массива. Следующий пример кода показывает, как выполнять итерацию по совпадениям для регулярного выражения `reg1`, чтобы создать строку для отображения в виде HTML-кода.</span><span class="sxs-lookup"><span data-stu-id="0f909-p119">Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.</span></span>

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

<span data-ttu-id="0f909-191">Ниже приведен пример правила `ItemHasKnownEntity`, которое указывает сущность `MeetingSuggestion` и регулярное выражение `CampSuggestion`.</span><span class="sxs-lookup"><span data-stu-id="0f909-191">The following is an example of an `ItemHasKnownEntity` rule that specifies the `MeetingSuggestion` entity and a regular expression named `CampSuggestion`.</span></span> <span data-ttu-id="0f909-192">Outlook активирует надстройку, если обнаруживает, что выбранный элемент содержит приглашение на собрание, а тема или текст содержит термин `WonderCamp`.</span><span class="sxs-lookup"><span data-stu-id="0f909-192">Outlook activates the add-in if it detects that the currently selected item contains a meeting suggestion, and the subject or body contains the term `WonderCamp`.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

<span data-ttu-id="0f909-193">В следующем примере кода используется метод `getFilteredEntitiesByName` текущего элемента, чтобы поместить в переменную `suggestions` массив обнаруженных приглашений на собрание для предыдущего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="0f909-193">The following code example uses `getFilteredEntitiesByName` on the current item to set a variable `suggestions` to an array of detected meeting suggestions for the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a><span data-ttu-id="0f909-194">См. также</span><span class="sxs-lookup"><span data-stu-id="0f909-194">See also</span></span>

- <span data-ttu-id="0f909-195">[Надстройка Outlook: номер заказа Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) — контекстная надстройка, которая активируется на основе соответствия регулярному выражению.</span><span class="sxs-lookup"><span data-stu-id="0f909-195">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - A sample contextual add-in that activates based on a regular expression match.</span></span>
- [<span data-ttu-id="0f909-196">Создание надстроек Outlook для форм чтения</span><span class="sxs-lookup"><span data-stu-id="0f909-196">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="0f909-197">Правила активации для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="0f909-197">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="0f909-198">Ограничения для активации и API JavaScript для надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="0f909-198">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="0f909-199">Сопоставление строк в элементе Outlook как известных сущностей</span><span class="sxs-lookup"><span data-stu-id="0f909-199">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
- [<span data-ttu-id="0f909-200">Рекомендации по использованию регулярных выражений в .NET Framework</span><span class="sxs-lookup"><span data-stu-id="0f909-200">Best Practices for Regular Expressions in the .NET Framework</span></span>](/dotnet/standard/base-types/best-practices)
