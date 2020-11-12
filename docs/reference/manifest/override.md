---
title: Элемент Override в файле манифеста
description: Элемент override позволяет указать значение параметра в зависимости от указанного условия.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996314"
---
# <a name="override-element"></a><span data-ttu-id="25fa5-103">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="25fa5-103">Override element</span></span>

<span data-ttu-id="25fa5-104">Предоставляет способ переопределения значения параметра манифеста в зависимости от указанного условия.</span><span class="sxs-lookup"><span data-stu-id="25fa5-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="25fa5-105">Существует два типа условий:</span><span class="sxs-lookup"><span data-stu-id="25fa5-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="25fa5-106">Языковой стандарт Office, отличный от используемого по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="25fa5-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="25fa5-107">Шаблон поддержки набора требований, отличный от шаблона по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="25fa5-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="25fa5-108">Существует два типа элементов: `<Override>` один — для переопределения языкового стандарта, который называется **локалетокеноверриде** , а другой — для переопределения набора требований, именуемого **рекуиременттокеноверриде**.</span><span class="sxs-lookup"><span data-stu-id="25fa5-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride** , and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="25fa5-109">Но `type` для элемента нет параметров `<Override>` .</span><span class="sxs-lookup"><span data-stu-id="25fa5-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="25fa5-110">Разница определяется родительским элементом и типом родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="25fa5-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="25fa5-111">`<Override>`Элемент, который находится внутри `<Token>` элемента `xsi:type` , который `RequirementToken` должен иметь тип **рекуиременттокеноверриде**.</span><span class="sxs-lookup"><span data-stu-id="25fa5-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="25fa5-112">`<Override>`Элемент внутри любого другого родительского элемента или внутри `<Override>` элемента типа `LocaleToken` должен иметь тип **локалетокеноверриде**.</span><span class="sxs-lookup"><span data-stu-id="25fa5-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="25fa5-113">Каждый тип описывается в отдельных разделах ниже.</span><span class="sxs-lookup"><span data-stu-id="25fa5-113">Each type is described in separate sections below.</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="25fa5-114">Элемент override элемента типа Локалетокеноверриде</span><span class="sxs-lookup"><span data-stu-id="25fa5-114">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="25fa5-115">`<Override>`Элемент выражает условное значение и может быть прочитано как "If... Then... " Оператор.</span><span class="sxs-lookup"><span data-stu-id="25fa5-115">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="25fa5-116">Если `<Override>` элемент имеет тип **локалетокеноверриде** , `Locale` атрибут является условием, а `Value` атрибут — консекуент.</span><span class="sxs-lookup"><span data-stu-id="25fa5-116">If the `<Override>` element is of type **LocaleTokenOverride** , then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="25fa5-117">Например, прочтите следующий текст: "при настройке языкового стандарта Office fr-FR отображается имя" Лектеур видéо "."</span><span class="sxs-lookup"><span data-stu-id="25fa5-117">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="25fa5-118">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="25fa5-118">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="25fa5-119">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="25fa5-119">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="25fa5-120">Содержится в</span><span class="sxs-lookup"><span data-stu-id="25fa5-120">Contained in</span></span>

|<span data-ttu-id="25fa5-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="25fa5-121">Element</span></span>|
|:-----|
|[<span data-ttu-id="25fa5-122">CitationText</span><span class="sxs-lookup"><span data-stu-id="25fa5-122">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="25fa5-123">Описание</span><span class="sxs-lookup"><span data-stu-id="25fa5-123">Description</span></span>](description.md)|
|[<span data-ttu-id="25fa5-124">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="25fa5-124">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="25fa5-125">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="25fa5-125">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="25fa5-126">DisplayName</span><span class="sxs-lookup"><span data-stu-id="25fa5-126">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="25fa5-127">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="25fa5-127">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="25fa5-128">IconUrl</span><span class="sxs-lookup"><span data-stu-id="25fa5-128">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="25fa5-129">QueryUri</span><span class="sxs-lookup"><span data-stu-id="25fa5-129">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="25fa5-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="25fa5-130">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="25fa5-131">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="25fa5-131">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="25fa5-132">Маркер</span><span class="sxs-lookup"><span data-stu-id="25fa5-132">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="25fa5-133">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="25fa5-133">Attributes</span></span>

|<span data-ttu-id="25fa5-134">Атрибут</span><span class="sxs-lookup"><span data-stu-id="25fa5-134">Attribute</span></span>|<span data-ttu-id="25fa5-135">Тип</span><span class="sxs-lookup"><span data-stu-id="25fa5-135">Type</span></span>|<span data-ttu-id="25fa5-136">Обязательный</span><span class="sxs-lookup"><span data-stu-id="25fa5-136">Required</span></span>|<span data-ttu-id="25fa5-137">Описание</span><span class="sxs-lookup"><span data-stu-id="25fa5-137">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="25fa5-138">Языковой стандарт</span><span class="sxs-lookup"><span data-stu-id="25fa5-138">Locale</span></span>|<span data-ttu-id="25fa5-139">string</span><span class="sxs-lookup"><span data-stu-id="25fa5-139">string</span></span>|<span data-ttu-id="25fa5-140">Обязательный</span><span class="sxs-lookup"><span data-stu-id="25fa5-140">required</span></span>|<span data-ttu-id="25fa5-141">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="25fa5-141">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="25fa5-142">Значение</span><span class="sxs-lookup"><span data-stu-id="25fa5-142">Value</span></span>|<span data-ttu-id="25fa5-143">string</span><span class="sxs-lookup"><span data-stu-id="25fa5-143">string</span></span>|<span data-ttu-id="25fa5-144">Обязательный</span><span class="sxs-lookup"><span data-stu-id="25fa5-144">required</span></span>|<span data-ttu-id="25fa5-145">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="25fa5-145">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="25fa5-146">Примеры</span><span class="sxs-lookup"><span data-stu-id="25fa5-146">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="25fa5-147">См. также</span><span class="sxs-lookup"><span data-stu-id="25fa5-147">See also</span></span>

- [<span data-ttu-id="25fa5-148">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="25fa5-148">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="25fa5-149">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="25fa5-149">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="25fa5-150">Элемент override элемента типа Рекуиременттокеноверриде</span><span class="sxs-lookup"><span data-stu-id="25fa5-150">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="25fa5-151">`<Override>`Элемент выражает условное значение и может быть прочитано как "If... Then... " Оператор.</span><span class="sxs-lookup"><span data-stu-id="25fa5-151">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="25fa5-152">Если `<Override>` элемент имеет тип **рекуиременттокеноверриде** , дочерний `<Requirements>` элемент выражает условие, а `Value` атрибут — консекуент.</span><span class="sxs-lookup"><span data-stu-id="25fa5-152">If the `<Override>` element is of type **RequirementTokenOverride** , then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="25fa5-153">Например, первое `<Override>` в приведенном ниже примере считывается, если текущая платформа поддерживает феатуреоне версии 1,7, а затем используйте строку "олдаддинверсион" вместо `${token.requirements}` маркера в URL-адресе в URL-адресе "бабушке" `<ExtendedOverrides>` (вместо строки по умолчанию "Upgrade"). "</span><span class="sxs-lookup"><span data-stu-id="25fa5-153">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="25fa5-154">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="25fa5-154">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="25fa5-155">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="25fa5-155">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="25fa5-156">Содержится в</span><span class="sxs-lookup"><span data-stu-id="25fa5-156">Contained in</span></span>

|<span data-ttu-id="25fa5-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="25fa5-157">Element</span></span>|
|:-----|
|[<span data-ttu-id="25fa5-158">Маркер</span><span class="sxs-lookup"><span data-stu-id="25fa5-158">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="25fa5-159">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="25fa5-159">Must contain</span></span>

|<span data-ttu-id="25fa5-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="25fa5-160">Element</span></span>|<span data-ttu-id="25fa5-161">Контентная</span><span class="sxs-lookup"><span data-stu-id="25fa5-161">Content</span></span>|<span data-ttu-id="25fa5-162">Почта</span><span class="sxs-lookup"><span data-stu-id="25fa5-162">Mail</span></span>|<span data-ttu-id="25fa5-163">Область задач</span><span class="sxs-lookup"><span data-stu-id="25fa5-163">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="25fa5-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="25fa5-164">Requirements</span></span>](requirements.md)|||<span data-ttu-id="25fa5-165">x</span><span class="sxs-lookup"><span data-stu-id="25fa5-165">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="25fa5-166">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="25fa5-166">Attributes</span></span>

|<span data-ttu-id="25fa5-167">Атрибут</span><span class="sxs-lookup"><span data-stu-id="25fa5-167">Attribute</span></span>|<span data-ttu-id="25fa5-168">Тип</span><span class="sxs-lookup"><span data-stu-id="25fa5-168">Type</span></span>|<span data-ttu-id="25fa5-169">Обязательный</span><span class="sxs-lookup"><span data-stu-id="25fa5-169">Required</span></span>|<span data-ttu-id="25fa5-170">Описание</span><span class="sxs-lookup"><span data-stu-id="25fa5-170">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="25fa5-171">Значение</span><span class="sxs-lookup"><span data-stu-id="25fa5-171">Value</span></span>|<span data-ttu-id="25fa5-172">string</span><span class="sxs-lookup"><span data-stu-id="25fa5-172">string</span></span>|<span data-ttu-id="25fa5-173">Обязательный</span><span class="sxs-lookup"><span data-stu-id="25fa5-173">required</span></span>|<span data-ttu-id="25fa5-174">Значение маркера "бабушке" при удовлетворении условия.</span><span class="sxs-lookup"><span data-stu-id="25fa5-174">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="25fa5-175">Пример</span><span class="sxs-lookup"><span data-stu-id="25fa5-175">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="25fa5-176">См. также</span><span class="sxs-lookup"><span data-stu-id="25fa5-176">See also</span></span>

- [<span data-ttu-id="25fa5-177">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="25fa5-177">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="25fa5-178">Указание элемента Requirements в манифесте</span><span class="sxs-lookup"><span data-stu-id="25fa5-178">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="25fa5-179">Сочетания клавиш</span><span class="sxs-lookup"><span data-stu-id="25fa5-179">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
