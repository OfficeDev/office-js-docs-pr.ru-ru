---
title: Элемент Override в файле манифеста
description: Элемент Переопределения позволяет указать значение параметра в зависимости от заданного условия.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: d2146cc1f44e829bc78076c8093b2ebf791dc722
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505341"
---
# <a name="override-element"></a><span data-ttu-id="5830d-103">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="5830d-103">Override element</span></span>

<span data-ttu-id="5830d-104">Предоставляет способ переопределения значения параметра манифеста в зависимости от указанного условия.</span><span class="sxs-lookup"><span data-stu-id="5830d-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="5830d-105">Существует два типа условий:</span><span class="sxs-lookup"><span data-stu-id="5830d-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="5830d-106">Локальный стандарт Office, который отличается от по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5830d-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="5830d-107">Шаблон поддержки набора требований, который отличается от шаблона по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5830d-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="5830d-108">Существует два типа элементов, один из них для переопределеть локаута, называемый `<Override>` **LocaleTokenOverride,** а другой — для переопределей набора требований, называемых **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="5830d-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride**, and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="5830d-109">Но параметра `type` для элемента `<Override>` нет.</span><span class="sxs-lookup"><span data-stu-id="5830d-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="5830d-110">Разница определяется родительским элементом и типом родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="5830d-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="5830d-111">Элемент, `<Override>` который находится внутри `<Token>` элемента, который является , должен быть `xsi:type` `RequirementToken` типа **RequirementTokenOverride**.</span><span class="sxs-lookup"><span data-stu-id="5830d-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="5830d-112">Элемент внутри любого другого родительского элемента или элемента типа должен быть типа `<Override>` `<Override>` `LocaleToken` **LocaleTokenOverride.**</span><span class="sxs-lookup"><span data-stu-id="5830d-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="5830d-113">Каждый тип описывается в отдельных разделах ниже.</span><span class="sxs-lookup"><span data-stu-id="5830d-113">Each type is described in separate sections below.</span></span> <span data-ttu-id="5830d-114">Дополнительные сведения об использовании этого элемента, когда он является ребенком элемента, см. в этой ссылке Работа с расширенными `<Token>` [переопределениями манифеста.](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="5830d-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="5830d-115">Переопределять элемент типа LocaleTokenOverride</span><span class="sxs-lookup"><span data-stu-id="5830d-115">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="5830d-116">Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление.</span><span class="sxs-lookup"><span data-stu-id="5830d-116">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5830d-117">Если элемент `<Override>` имеет тип **LocaleTokenOverride,** то атрибут является условием, а атрибут `Locale` — `Value` последующим.</span><span class="sxs-lookup"><span data-stu-id="5830d-117">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="5830d-118">Например, ниже приводится следующий текст: "Если параметр office locale является fr-fr, то имя отображения — "Lecteur vidéo".</span><span class="sxs-lookup"><span data-stu-id="5830d-118">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="5830d-119">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="5830d-119">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="5830d-120">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="5830d-120">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="5830d-121">Содержится в</span><span class="sxs-lookup"><span data-stu-id="5830d-121">Contained in</span></span>

|<span data-ttu-id="5830d-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="5830d-122">Element</span></span>|
|:-----|
|[<span data-ttu-id="5830d-123">CitationText</span><span class="sxs-lookup"><span data-stu-id="5830d-123">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="5830d-124">Описание</span><span class="sxs-lookup"><span data-stu-id="5830d-124">Description</span></span>](description.md)|
|[<span data-ttu-id="5830d-125">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="5830d-125">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="5830d-126">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="5830d-126">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="5830d-127">DisplayName</span><span class="sxs-lookup"><span data-stu-id="5830d-127">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="5830d-128">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="5830d-128">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="5830d-129">IconUrl</span><span class="sxs-lookup"><span data-stu-id="5830d-129">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="5830d-130">QueryUri</span><span class="sxs-lookup"><span data-stu-id="5830d-130">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="5830d-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5830d-131">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="5830d-132">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="5830d-132">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="5830d-133">Маркер</span><span class="sxs-lookup"><span data-stu-id="5830d-133">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="5830d-134">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5830d-134">Attributes</span></span>

|<span data-ttu-id="5830d-135">Атрибут</span><span class="sxs-lookup"><span data-stu-id="5830d-135">Attribute</span></span>|<span data-ttu-id="5830d-136">Тип</span><span class="sxs-lookup"><span data-stu-id="5830d-136">Type</span></span>|<span data-ttu-id="5830d-137">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5830d-137">Required</span></span>|<span data-ttu-id="5830d-138">Описание</span><span class="sxs-lookup"><span data-stu-id="5830d-138">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5830d-139">Языковой стандарт</span><span class="sxs-lookup"><span data-stu-id="5830d-139">Locale</span></span>|<span data-ttu-id="5830d-140">string</span><span class="sxs-lookup"><span data-stu-id="5830d-140">string</span></span>|<span data-ttu-id="5830d-141">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5830d-141">required</span></span>|<span data-ttu-id="5830d-142">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="5830d-142">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="5830d-143">Значение</span><span class="sxs-lookup"><span data-stu-id="5830d-143">Value</span></span>|<span data-ttu-id="5830d-144">string</span><span class="sxs-lookup"><span data-stu-id="5830d-144">string</span></span>|<span data-ttu-id="5830d-145">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5830d-145">required</span></span>|<span data-ttu-id="5830d-146">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="5830d-146">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="5830d-147">Примеры</span><span class="sxs-lookup"><span data-stu-id="5830d-147">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="5830d-148">См. также</span><span class="sxs-lookup"><span data-stu-id="5830d-148">See also</span></span>

- [<span data-ttu-id="5830d-149">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="5830d-149">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="5830d-150">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="5830d-150">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="5830d-151">Переопределять элемент типа RequirementTokenOverride</span><span class="sxs-lookup"><span data-stu-id="5830d-151">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="5830d-152">Элемент `<Override>` выражает условный и может быть прочитано как "Если ... затем ..." заявление.</span><span class="sxs-lookup"><span data-stu-id="5830d-152">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5830d-153">Если элемент `<Override>` имеет тип **RequirementTokenOverride,** то детский элемент выражает условие, а атрибут — `<Requirements>` `Value` следовательно.</span><span class="sxs-lookup"><span data-stu-id="5830d-153">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="5830d-154">Например, первое из следующих строк гласит: "Если текущая платформа поддерживает `<Override>` версию FeatureOne 1.7, используйте строку "oldAddinVersion" вместо маркера в URL-адресе бабушки и дедушки (вместо строки по умолчанию `${token.requirements}` `<ExtendedOverrides>` "обновление") ".</span><span class="sxs-lookup"><span data-stu-id="5830d-154">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="5830d-155">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="5830d-155">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="5830d-156">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="5830d-156">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="5830d-157">Содержится в</span><span class="sxs-lookup"><span data-stu-id="5830d-157">Contained in</span></span>

|<span data-ttu-id="5830d-158">Элемент</span><span class="sxs-lookup"><span data-stu-id="5830d-158">Element</span></span>|
|:-----|
|[<span data-ttu-id="5830d-159">Маркер</span><span class="sxs-lookup"><span data-stu-id="5830d-159">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="5830d-160">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="5830d-160">Must contain</span></span>

|<span data-ttu-id="5830d-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="5830d-161">Element</span></span>|<span data-ttu-id="5830d-162">Контентная</span><span class="sxs-lookup"><span data-stu-id="5830d-162">Content</span></span>|<span data-ttu-id="5830d-163">Почта</span><span class="sxs-lookup"><span data-stu-id="5830d-163">Mail</span></span>|<span data-ttu-id="5830d-164">Область задач</span><span class="sxs-lookup"><span data-stu-id="5830d-164">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5830d-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="5830d-165">Requirements</span></span>](requirements.md)|||<span data-ttu-id="5830d-166">x</span><span class="sxs-lookup"><span data-stu-id="5830d-166">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="5830d-167">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5830d-167">Attributes</span></span>

|<span data-ttu-id="5830d-168">Атрибут</span><span class="sxs-lookup"><span data-stu-id="5830d-168">Attribute</span></span>|<span data-ttu-id="5830d-169">Тип</span><span class="sxs-lookup"><span data-stu-id="5830d-169">Type</span></span>|<span data-ttu-id="5830d-170">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5830d-170">Required</span></span>|<span data-ttu-id="5830d-171">Описание</span><span class="sxs-lookup"><span data-stu-id="5830d-171">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5830d-172">Значение</span><span class="sxs-lookup"><span data-stu-id="5830d-172">Value</span></span>|<span data-ttu-id="5830d-173">string</span><span class="sxs-lookup"><span data-stu-id="5830d-173">string</span></span>|<span data-ttu-id="5830d-174">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5830d-174">required</span></span>|<span data-ttu-id="5830d-175">Значение маркера дедушек и дедушек при условии удовлетворены.</span><span class="sxs-lookup"><span data-stu-id="5830d-175">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="5830d-176">Пример</span><span class="sxs-lookup"><span data-stu-id="5830d-176">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="5830d-177">См. также</span><span class="sxs-lookup"><span data-stu-id="5830d-177">See also</span></span>

- [<span data-ttu-id="5830d-178">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="5830d-178">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="5830d-179">Указание элемента Requirements в манифесте</span><span class="sxs-lookup"><span data-stu-id="5830d-179">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="5830d-180">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="5830d-180">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
