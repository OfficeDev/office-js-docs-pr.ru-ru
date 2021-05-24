---
title: Элемент Override в файле манифеста
description: Элемент Переопределения позволяет указать значение параметра в зависимости от заданного условия.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd270fa19750810238b42c26c2abc35a61c1bac8
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590906"
---
# <a name="override-element"></a><span data-ttu-id="2f5e5-103">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="2f5e5-103">Override element</span></span>

<span data-ttu-id="2f5e5-104">Предоставляет способ переопределения значения параметра манифеста в зависимости от указанного условия.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="2f5e5-105">Существует три типа условий:</span><span class="sxs-lookup"><span data-stu-id="2f5e5-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="2f5e5-106">Локальный Office, который отличается от по `LocaleToken` умолчанию, называется **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="2f5e5-107">Шаблон поддержки набора требований, который отличается от шаблона по `RequirementToken` умолчанию, называемого **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="2f5e5-108">Источник отличается от по `Runtime` умолчанию, называется **RuntimeOverride**.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-108">The source is different from the default `Runtime`, called **RuntimeOverride**.</span></span>

<span data-ttu-id="2f5e5-109">Элемент, который находится внутри элемента, должен `<Override>` иметь тип `<Runtime>` **RuntimeOverride.**</span><span class="sxs-lookup"><span data-stu-id="2f5e5-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="2f5e5-110">Атрибут элемента `overrideType` не `<Override>` существует.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="2f5e5-111">Разница определяется родительским элементом и типом родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="2f5e5-112">Элемент, `<Override>` который находится внутри `<Token>` элемента, который является , должен быть `xsi:type` `RequirementToken` типа **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="2f5e5-113">Элемент внутри любого другого родительского элемента или элемента типа должен быть типа `<Override>` `<Override>` `LocaleToken` **LocaleTokenOverride.**</span><span class="sxs-lookup"><span data-stu-id="2f5e5-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="2f5e5-114">Дополнительные сведения об использовании этого элемента, когда он является ребенком элемента, см. в этой ссылке Работа с расширенными `<Token>` [переопределениями манифеста.](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="2f5e5-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="2f5e5-115">Каждый тип описан в отдельных разделах позднее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="2f5e5-116">Элемент Переопределения для `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="2f5e5-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="2f5e5-117">Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="2f5e5-118">Если элемент `<Override>` имеет тип **LocaleTokenOverride,** то атрибут является условием, а атрибут `Locale` — `Value` последующим.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="2f5e5-119">Например, ниже приводится следующий текст: "Если параметр Office fr-fr, то имя отображения — "Lecteur vidéo".</span><span class="sxs-lookup"><span data-stu-id="2f5e5-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="2f5e5-120">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2f5e5-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="2f5e5-121">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2f5e5-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="2f5e5-122">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2f5e5-122">Contained in</span></span>

|<span data-ttu-id="2f5e5-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="2f5e5-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="2f5e5-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="2f5e5-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="2f5e5-125">Описание</span><span class="sxs-lookup"><span data-stu-id="2f5e5-125">Description</span></span>](description.md)|
|[<span data-ttu-id="2f5e5-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="2f5e5-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="2f5e5-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="2f5e5-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="2f5e5-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="2f5e5-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="2f5e5-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="2f5e5-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="2f5e5-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="2f5e5-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="2f5e5-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="2f5e5-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="2f5e5-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="2f5e5-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="2f5e5-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="2f5e5-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="2f5e5-134">Маркер</span><span class="sxs-lookup"><span data-stu-id="2f5e5-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="2f5e5-135">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2f5e5-135">Attributes</span></span>

|<span data-ttu-id="2f5e5-136">Атрибут</span><span class="sxs-lookup"><span data-stu-id="2f5e5-136">Attribute</span></span>|<span data-ttu-id="2f5e5-137">Тип</span><span class="sxs-lookup"><span data-stu-id="2f5e5-137">Type</span></span>|<span data-ttu-id="2f5e5-138">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f5e5-138">Required</span></span>|<span data-ttu-id="2f5e5-139">Описание</span><span class="sxs-lookup"><span data-stu-id="2f5e5-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2f5e5-140">Языковой стандарт</span><span class="sxs-lookup"><span data-stu-id="2f5e5-140">Locale</span></span>|<span data-ttu-id="2f5e5-141">string</span><span class="sxs-lookup"><span data-stu-id="2f5e5-141">string</span></span>|<span data-ttu-id="2f5e5-142">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f5e5-142">required</span></span>|<span data-ttu-id="2f5e5-143">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="2f5e5-144">Значение</span><span class="sxs-lookup"><span data-stu-id="2f5e5-144">Value</span></span>|<span data-ttu-id="2f5e5-145">string</span><span class="sxs-lookup"><span data-stu-id="2f5e5-145">string</span></span>|<span data-ttu-id="2f5e5-146">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f5e5-146">required</span></span>|<span data-ttu-id="2f5e5-147">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="2f5e5-148">Примеры</span><span class="sxs-lookup"><span data-stu-id="2f5e5-148">Examples</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a><span data-ttu-id="2f5e5-149">См. также</span><span class="sxs-lookup"><span data-stu-id="2f5e5-149">See also</span></span>

- [<span data-ttu-id="2f5e5-150">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="2f5e5-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="2f5e5-151">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="2f5e5-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="2f5e5-152">Элемент Переопределения для `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="2f5e5-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="2f5e5-153">Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="2f5e5-154">Если элемент `<Override>` имеет тип **RequirementTokenOverride,** то детский элемент выражает условие, а атрибут — `<Requirements>` `Value` следовательно.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="2f5e5-155">Например, первое из следующих строк гласит: "Если текущая платформа поддерживает `<Override>` версию FeatureOne 1.7, используйте строку "oldAddinVersion" вместо маркера в URL-адресе бабушки и дедушки (вместо строки по умолчанию `${token.requirements}` `<ExtendedOverrides>` "обновление") ".</span><span class="sxs-lookup"><span data-stu-id="2f5e5-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

<span data-ttu-id="2f5e5-156">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2f5e5-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="2f5e5-157">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2f5e5-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="2f5e5-158">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2f5e5-158">Contained in</span></span>

|<span data-ttu-id="2f5e5-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="2f5e5-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="2f5e5-160">Маркер</span><span class="sxs-lookup"><span data-stu-id="2f5e5-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="2f5e5-161">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="2f5e5-161">Must contain</span></span>

|<span data-ttu-id="2f5e5-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="2f5e5-162">Element</span></span>|<span data-ttu-id="2f5e5-163">Контентная</span><span class="sxs-lookup"><span data-stu-id="2f5e5-163">Content</span></span>|<span data-ttu-id="2f5e5-164">Почта</span><span class="sxs-lookup"><span data-stu-id="2f5e5-164">Mail</span></span>|<span data-ttu-id="2f5e5-165">Область задач</span><span class="sxs-lookup"><span data-stu-id="2f5e5-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2f5e5-166">Requirements</span><span class="sxs-lookup"><span data-stu-id="2f5e5-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="2f5e5-167">x</span><span class="sxs-lookup"><span data-stu-id="2f5e5-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="2f5e5-168">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2f5e5-168">Attributes</span></span>

|<span data-ttu-id="2f5e5-169">Атрибут</span><span class="sxs-lookup"><span data-stu-id="2f5e5-169">Attribute</span></span>|<span data-ttu-id="2f5e5-170">Тип</span><span class="sxs-lookup"><span data-stu-id="2f5e5-170">Type</span></span>|<span data-ttu-id="2f5e5-171">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f5e5-171">Required</span></span>|<span data-ttu-id="2f5e5-172">Описание</span><span class="sxs-lookup"><span data-stu-id="2f5e5-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2f5e5-173">Значение</span><span class="sxs-lookup"><span data-stu-id="2f5e5-173">Value</span></span>|<span data-ttu-id="2f5e5-174">string</span><span class="sxs-lookup"><span data-stu-id="2f5e5-174">string</span></span>|<span data-ttu-id="2f5e5-175">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f5e5-175">required</span></span>|<span data-ttu-id="2f5e5-176">Значение маркера дедушек и дедушек при условии удовлетворены.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="2f5e5-177">Пример</span><span class="sxs-lookup"><span data-stu-id="2f5e5-177">Example</span></span>

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a><span data-ttu-id="2f5e5-178">См. также</span><span class="sxs-lookup"><span data-stu-id="2f5e5-178">See also</span></span>

- [<span data-ttu-id="2f5e5-179">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="2f5e5-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="2f5e5-180">Указание элемента Requirements в манифесте</span><span class="sxs-lookup"><span data-stu-id="2f5e5-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="2f5e5-181">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="2f5e5-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a><span data-ttu-id="2f5e5-182">Элемент Переопределения для `Runtime`</span><span class="sxs-lookup"><span data-stu-id="2f5e5-182">Override element for `Runtime`</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f5e5-183">Поддержка этого элемента была представлена в наборе требований к почтовым ящикам [1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) с функцией активации на основе [событий.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="2f5e5-183">Support for this element was introduced in [Mailbox requirement set 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) with the [event-based activation feature](../../outlook/autolaunch.md).</span></span> <span data-ttu-id="2f5e5-184">См [клиенты и платформы](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-184">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="2f5e5-185">Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-185">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="2f5e5-186">Если элемент `<Override>` имеет тип **RuntimeOverride,** то атрибут является условием, а атрибут `type` — `resid` последующим.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-186">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="2f5e5-187">Например, ниже приводится следующее: "Если тип является "javascript", то это `resid` "JSRuntime.Url". Outlook Этот элемент требуется для обработчиков [точеки расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)</span><span class="sxs-lookup"><span data-stu-id="2f5e5-187">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="2f5e5-188">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="2f5e5-188">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="2f5e5-189">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2f5e5-189">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="2f5e5-190">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2f5e5-190">Contained in</span></span>

- [<span data-ttu-id="2f5e5-191">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="2f5e5-191">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="2f5e5-192">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2f5e5-192">Attributes</span></span>

|<span data-ttu-id="2f5e5-193">Атрибут</span><span class="sxs-lookup"><span data-stu-id="2f5e5-193">Attribute</span></span>|<span data-ttu-id="2f5e5-194">Тип</span><span class="sxs-lookup"><span data-stu-id="2f5e5-194">Type</span></span>|<span data-ttu-id="2f5e5-195">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f5e5-195">Required</span></span>|<span data-ttu-id="2f5e5-196">Описание</span><span class="sxs-lookup"><span data-stu-id="2f5e5-196">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2f5e5-197">**type**</span><span class="sxs-lookup"><span data-stu-id="2f5e5-197">**type**</span></span>|<span data-ttu-id="2f5e5-198">string</span><span class="sxs-lookup"><span data-stu-id="2f5e5-198">string</span></span>|<span data-ttu-id="2f5e5-199">Да</span><span class="sxs-lookup"><span data-stu-id="2f5e5-199">Yes</span></span>|<span data-ttu-id="2f5e5-200">Указывает язык для этого переопределения.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-200">Specifies the language for this override.</span></span> <span data-ttu-id="2f5e5-201">В настоящее `"javascript"` время это единственный поддерживаемый вариант.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-201">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="2f5e5-202">**resid**</span><span class="sxs-lookup"><span data-stu-id="2f5e5-202">**resid**</span></span>|<span data-ttu-id="2f5e5-203">string</span><span class="sxs-lookup"><span data-stu-id="2f5e5-203">string</span></span>|<span data-ttu-id="2f5e5-204">Да</span><span class="sxs-lookup"><span data-stu-id="2f5e5-204">Yes</span></span>|<span data-ttu-id="2f5e5-205">Указывает расположение URL-адреса файла JavaScript, который должен переопределять расположение URL-адреса HTML по умолчанию, определенного в родительском элементе [Runtime.](runtime.md) `resid`</span><span class="sxs-lookup"><span data-stu-id="2f5e5-205">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="2f5e5-206">Символ может быть не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента `Resources` элемента.</span><span class="sxs-lookup"><span data-stu-id="2f5e5-206">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="2f5e5-207">Примеры</span><span class="sxs-lookup"><span data-stu-id="2f5e5-207">Examples</span></span>

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a><span data-ttu-id="2f5e5-208">См. также</span><span class="sxs-lookup"><span data-stu-id="2f5e5-208">See also</span></span>

- [<span data-ttu-id="2f5e5-209">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="2f5e5-209">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="2f5e5-210">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="2f5e5-210">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
