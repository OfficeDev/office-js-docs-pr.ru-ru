---
title: Элемент Override в файле манифеста
description: Элемент Override позволяет указать значение параметра в зависимости от заданного состояния.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 131d72883d050038e2df5b7d8bbca033af9e6ee4
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555159"
---
# <a name="override-element"></a><span data-ttu-id="e9550-103">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="e9550-103">Override element</span></span>

<span data-ttu-id="e9550-104">Предоставляет способ переопределить значение параметра манифеста в зависимости от заданного состояния.</span><span class="sxs-lookup"><span data-stu-id="e9550-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="e9550-105">Существует три вида условий:</span><span class="sxs-lookup"><span data-stu-id="e9550-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="e9550-106">В Office, который отличается от по `LocaleToken` умолчанию, называется **LocaleTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="e9550-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="e9550-107">Шаблон поддержки набора требований, который отличается от шаблона по `RequirementToken` умолчанию, **называемого RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="e9550-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="e9550-108">Источник отличается от по `Runtime` умолчанию, называется **RuntimeOverride (в настоящее** время в предварительном просмотре).</span><span class="sxs-lookup"><span data-stu-id="e9550-108">The source is different from the default `Runtime`, called **RuntimeOverride** (currently in preview).</span></span>

<span data-ttu-id="e9550-109">Элемент, `<Override>` который находится внутри `<Runtime>` элемента, должен быть типа **RuntimeOverride.**</span><span class="sxs-lookup"><span data-stu-id="e9550-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="e9550-110">Атрибут элемента `overrideType` не `<Override>` существует.</span><span class="sxs-lookup"><span data-stu-id="e9550-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="e9550-111">Разница определяется родительским элементом и типом родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="e9550-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="e9550-112">Элемент, `<Override>` который находится внутри `<Token>` элемента, который `xsi:type` `RequirementToken` является, должен быть типа **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="e9550-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="e9550-113">Элемент `<Override>` внутри любого другого родительского элемента, или `<Override>` внутри элемента `LocaleToken` типа, должен быть типа **LocaleTokenOverride.**</span><span class="sxs-lookup"><span data-stu-id="e9550-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="e9550-114">Для получения дополнительной информации об использовании этого элемента, когда он является ребенком `<Token>` элемента, см [Работа с расширенными переопределениями манифеста.](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="e9550-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="e9550-115">Каждый тип описан в отдельных разделах позже в этой статье.</span><span class="sxs-lookup"><span data-stu-id="e9550-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="e9550-116">Элемент переопределения для `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="e9550-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="e9550-117">Элемент `<Override>` выражает условный и может быть прочитан как "Если ... затем ..." утверждение.</span><span class="sxs-lookup"><span data-stu-id="e9550-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="e9550-118">Если `<Override>` элемент типа **LocaleTokenOverride**, `Locale` то атрибут является условием, и атрибут является `Value` последующим.</span><span class="sxs-lookup"><span data-stu-id="e9550-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="e9550-119">Например, ниже приводится следующее: "Если Office настройки является fr-fr, то имя дисплея -" Lecteur vid'o".</span><span class="sxs-lookup"><span data-stu-id="e9550-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="e9550-120">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="e9550-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="e9550-121">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e9550-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="e9550-122">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e9550-122">Contained in</span></span>

|<span data-ttu-id="e9550-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9550-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="e9550-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="e9550-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="e9550-125">Описание</span><span class="sxs-lookup"><span data-stu-id="e9550-125">Description</span></span>](description.md)|
|[<span data-ttu-id="e9550-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="e9550-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="e9550-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="e9550-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="e9550-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="e9550-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="e9550-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="e9550-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="e9550-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="e9550-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="e9550-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="e9550-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="e9550-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e9550-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="e9550-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="e9550-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="e9550-134">Маркер</span><span class="sxs-lookup"><span data-stu-id="e9550-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="e9550-135">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e9550-135">Attributes</span></span>

|<span data-ttu-id="e9550-136">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e9550-136">Attribute</span></span>|<span data-ttu-id="e9550-137">Тип</span><span class="sxs-lookup"><span data-stu-id="e9550-137">Type</span></span>|<span data-ttu-id="e9550-138">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9550-138">Required</span></span>|<span data-ttu-id="e9550-139">Описание</span><span class="sxs-lookup"><span data-stu-id="e9550-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="e9550-140">Языковой стандарт</span><span class="sxs-lookup"><span data-stu-id="e9550-140">Locale</span></span>|<span data-ttu-id="e9550-141">string</span><span class="sxs-lookup"><span data-stu-id="e9550-141">string</span></span>|<span data-ttu-id="e9550-142">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9550-142">required</span></span>|<span data-ttu-id="e9550-143">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="e9550-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="e9550-144">Значение</span><span class="sxs-lookup"><span data-stu-id="e9550-144">Value</span></span>|<span data-ttu-id="e9550-145">string</span><span class="sxs-lookup"><span data-stu-id="e9550-145">string</span></span>|<span data-ttu-id="e9550-146">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9550-146">required</span></span>|<span data-ttu-id="e9550-147">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="e9550-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="e9550-148">Примеры</span><span class="sxs-lookup"><span data-stu-id="e9550-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="e9550-149">См. также</span><span class="sxs-lookup"><span data-stu-id="e9550-149">See also</span></span>

- [<span data-ttu-id="e9550-150">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="e9550-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="e9550-151">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="e9550-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="e9550-152">Элемент переопределения для `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="e9550-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="e9550-153">Элемент `<Override>` выражает условный и может быть прочитан как "Если ... затем ..." утверждение.</span><span class="sxs-lookup"><span data-stu-id="e9550-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="e9550-154">Если `<Override>` элемент типа **RequirementTokenOverride**, то `<Requirements>` элемент ребенка выражает условие, и атрибут является `Value` последующим.</span><span class="sxs-lookup"><span data-stu-id="e9550-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="e9550-155">Например, первый из `<Override>` следующих строк читается: "Если текущая платформа поддерживает версию FeatureOne 1.7, то используйте строку 'oldAddinVersion' вместо `${token.requirements}` маркера в URL дедушки и дедушки `<ExtendedOverrides>` (вместо строки по умолчанию 'обновление')".</span><span class="sxs-lookup"><span data-stu-id="e9550-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="e9550-156">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="e9550-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="e9550-157">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e9550-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="e9550-158">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e9550-158">Contained in</span></span>

|<span data-ttu-id="e9550-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9550-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="e9550-160">Маркер</span><span class="sxs-lookup"><span data-stu-id="e9550-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="e9550-161">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="e9550-161">Must contain</span></span>

|<span data-ttu-id="e9550-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="e9550-162">Element</span></span>|<span data-ttu-id="e9550-163">Контентная</span><span class="sxs-lookup"><span data-stu-id="e9550-163">Content</span></span>|<span data-ttu-id="e9550-164">Почта</span><span class="sxs-lookup"><span data-stu-id="e9550-164">Mail</span></span>|<span data-ttu-id="e9550-165">Область задач</span><span class="sxs-lookup"><span data-stu-id="e9550-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e9550-166">Requirements</span><span class="sxs-lookup"><span data-stu-id="e9550-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="e9550-167">x</span><span class="sxs-lookup"><span data-stu-id="e9550-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="e9550-168">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e9550-168">Attributes</span></span>

|<span data-ttu-id="e9550-169">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e9550-169">Attribute</span></span>|<span data-ttu-id="e9550-170">Тип</span><span class="sxs-lookup"><span data-stu-id="e9550-170">Type</span></span>|<span data-ttu-id="e9550-171">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9550-171">Required</span></span>|<span data-ttu-id="e9550-172">Описание</span><span class="sxs-lookup"><span data-stu-id="e9550-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="e9550-173">Значение</span><span class="sxs-lookup"><span data-stu-id="e9550-173">Value</span></span>|<span data-ttu-id="e9550-174">string</span><span class="sxs-lookup"><span data-stu-id="e9550-174">string</span></span>|<span data-ttu-id="e9550-175">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9550-175">required</span></span>|<span data-ttu-id="e9550-176">Значение знака бабушки и дедушки, когда условие удовлетворено.</span><span class="sxs-lookup"><span data-stu-id="e9550-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="e9550-177">Пример</span><span class="sxs-lookup"><span data-stu-id="e9550-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="e9550-178">См. также</span><span class="sxs-lookup"><span data-stu-id="e9550-178">See also</span></span>

- [<span data-ttu-id="e9550-179">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="e9550-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="e9550-180">Указание элемента Requirements в манифесте</span><span class="sxs-lookup"><span data-stu-id="e9550-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="e9550-181">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="e9550-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime-preview"></a><span data-ttu-id="e9550-182">Элемент переопределения `Runtime` для (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="e9550-182">Override element for `Runtime` (preview)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e9550-183">Эта функция поддерживается только для [предварительного](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) просмотра Outlook веб-сайтах и Windows с Microsoft 365 подпиской.</span><span class="sxs-lookup"><span data-stu-id="e9550-183">This feature is only supported for [preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="e9550-184">Для получения более подробной [информации см Outlook.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="e9550-184">For more details, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>
>
> <span data-ttu-id="e9550-185">Поскольку функции предварительного просмотра могут быть изменения без предварительного уведомления, они не должны использоваться в производственных дополнениях.</span><span class="sxs-lookup"><span data-stu-id="e9550-185">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

<span data-ttu-id="e9550-186">Элемент `<Override>` выражает условный и может быть прочитан как "Если ... затем ..." утверждение.</span><span class="sxs-lookup"><span data-stu-id="e9550-186">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="e9550-187">Если `<Override>` элемент типа **RuntimeOverride**, то `type` атрибут является условием, и `resid` атрибут является последующим.</span><span class="sxs-lookup"><span data-stu-id="e9550-187">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="e9550-188">Например, ниже приводится следующее: "Если тип "JavaScript", `resid` то 'JSRuntime.Url'." Outlook Рабочий стол требует этого элемента [для обработчиков токов точки расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)</span><span class="sxs-lookup"><span data-stu-id="e9550-188">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="e9550-189">**Тип надстройки:** почтовая</span><span class="sxs-lookup"><span data-stu-id="e9550-189">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="e9550-190">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e9550-190">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="e9550-191">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e9550-191">Contained in</span></span>

- [<span data-ttu-id="e9550-192">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="e9550-192">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="e9550-193">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e9550-193">Attributes</span></span>

|<span data-ttu-id="e9550-194">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e9550-194">Attribute</span></span>|<span data-ttu-id="e9550-195">Тип</span><span class="sxs-lookup"><span data-stu-id="e9550-195">Type</span></span>|<span data-ttu-id="e9550-196">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e9550-196">Required</span></span>|<span data-ttu-id="e9550-197">Описание</span><span class="sxs-lookup"><span data-stu-id="e9550-197">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="e9550-198">**type**</span><span class="sxs-lookup"><span data-stu-id="e9550-198">**type**</span></span>|<span data-ttu-id="e9550-199">string</span><span class="sxs-lookup"><span data-stu-id="e9550-199">string</span></span>|<span data-ttu-id="e9550-200">Да</span><span class="sxs-lookup"><span data-stu-id="e9550-200">Yes</span></span>|<span data-ttu-id="e9550-201">Определяет язык для этого переопределения.</span><span class="sxs-lookup"><span data-stu-id="e9550-201">Specifies the language for this override.</span></span> <span data-ttu-id="e9550-202">В настоящее `"javascript"` время это единственный поддерживаемый вариант.</span><span class="sxs-lookup"><span data-stu-id="e9550-202">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="e9550-203">**resid**</span><span class="sxs-lookup"><span data-stu-id="e9550-203">**resid**</span></span>|<span data-ttu-id="e9550-204">string</span><span class="sxs-lookup"><span data-stu-id="e9550-204">string</span></span>|<span data-ttu-id="e9550-205">Да</span><span class="sxs-lookup"><span data-stu-id="e9550-205">Yes</span></span>|<span data-ttu-id="e9550-206">Определяется местоположение URL-адреса файла JavaScript, который должен переопределить местоположение URL HTML по умолчанию, определяемого в [родительском](runtime.md) элементе `resid` Runtime.</span><span class="sxs-lookup"><span data-stu-id="e9550-206">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="e9550-207">Может `resid` быть не более 32 символов и должен `id` соответствовать атрибуту `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="e9550-207">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="e9550-208">Примеры</span><span class="sxs-lookup"><span data-stu-id="e9550-208">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="e9550-209">См. также</span><span class="sxs-lookup"><span data-stu-id="e9550-209">See also</span></span>

- [<span data-ttu-id="e9550-210">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="e9550-210">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="e9550-211">Настройте Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="e9550-211">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
