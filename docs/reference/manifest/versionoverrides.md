---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация элемента VersionOverrides для Office дополнительных дополнительных виленок (XML).
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 0a70ded82b4603b1ac70698947a4710a4a44b5b6
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555152"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="bffce-103">Элемент VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="bffce-103">VersionOverrides element</span></span>

<span data-ttu-id="bffce-p101">Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.</span><span class="sxs-lookup"><span data-stu-id="bffce-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="bffce-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bffce-107">Attributes</span></span>

|  <span data-ttu-id="bffce-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="bffce-108">Attribute</span></span>  |  <span data-ttu-id="bffce-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bffce-109">Required</span></span>  |  <span data-ttu-id="bffce-110">Описание</span><span class="sxs-lookup"><span data-stu-id="bffce-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bffce-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="bffce-111">**xmlns**</span></span>       |  <span data-ttu-id="bffce-112">Да</span><span class="sxs-lookup"><span data-stu-id="bffce-112">Yes</span></span>  |  <span data-ttu-id="bffce-113">ВерсияОвергорайды схема пространства имен.</span><span class="sxs-lookup"><span data-stu-id="bffce-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="bffce-114">Разрешенные значения варьируются в `<VersionOverrides>` зависимости от **значения xsi:типа** этого элемента **и значения xsi:типа** родительского `<OfficeApp>` элемента.</span><span class="sxs-lookup"><span data-stu-id="bffce-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="bffce-115">Ниже [приведены значения пространства имен.](#namespace-values)</span><span class="sxs-lookup"><span data-stu-id="bffce-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="bffce-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="bffce-116">**xsi:type**</span></span>  |  <span data-ttu-id="bffce-117">Да</span><span class="sxs-lookup"><span data-stu-id="bffce-117">Yes</span></span>  | <span data-ttu-id="bffce-p103">Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="bffce-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="bffce-120">Значения пространства имен</span><span class="sxs-lookup"><span data-stu-id="bffce-120">Namespace values</span></span>

<span data-ttu-id="bffce-121">Ниже приводится перечне требуемое значение **значения xmlns** в зависимости **от значения xsi:type** родительского `<OfficeApp>` элемента.</span><span class="sxs-lookup"><span data-stu-id="bffce-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="bffce-122">**TaskPaneApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны `http://schemas.microsoft.com/office/taskpaneappversionoverrides` быть.</span><span class="sxs-lookup"><span data-stu-id="bffce-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="bffce-123">**ContentApp** поддерживает только версию 1.0 VersionOverrides, и **xmlns** должны `http://schemas.microsoft.com/office/contentappversionoverrides` быть.</span><span class="sxs-lookup"><span data-stu-id="bffce-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="bffce-124">**MailApp** поддерживает версии 1.0 и 1.1 VersionOverrides, поэтому **значение xmlns варьируется** в `<VersionOverrides>` зависимости от **значения xsi:type** этого элемента:</span><span class="sxs-lookup"><span data-stu-id="bffce-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="bffce-125">Когда **xsi:type** `VersionOverridesV1_0` есть, **xmlns** должен `http://schemas.microsoft.com/office/mailappversionoverrides` быть.</span><span class="sxs-lookup"><span data-stu-id="bffce-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="bffce-126">Когда **xsi:type** `VersionOverridesV1_1` есть, **xmlns** должен `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` быть.</span><span class="sxs-lookup"><span data-stu-id="bffce-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="bffce-127">В настоящее Outlook 2016 или позже поддерживает схему VersionOverrides v1.1 и `VersionOverridesV1_1` тип.</span><span class="sxs-lookup"><span data-stu-id="bffce-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bffce-128">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bffce-128">Child elements</span></span>

|  <span data-ttu-id="bffce-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="bffce-129">Element</span></span> |  <span data-ttu-id="bffce-130">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bffce-130">Required</span></span>  |  <span data-ttu-id="bffce-131">Описание</span><span class="sxs-lookup"><span data-stu-id="bffce-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bffce-132">**Описание**</span><span class="sxs-lookup"><span data-stu-id="bffce-132">**Description**</span></span>    |  <span data-ttu-id="bffce-133">Нет</span><span class="sxs-lookup"><span data-stu-id="bffce-133">No</span></span>   |  <span data-ttu-id="bffce-134">Описывает надстройку.</span><span class="sxs-lookup"><span data-stu-id="bffce-134">Describes the add-in.</span></span> <span data-ttu-id="bffce-135">Переопределяет элемент `Description` в любой родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="bffce-135">This overrides the `Description` element in any parent portion of the manifest.</span></span> <span data-ttu-id="bffce-136">Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bffce-136">The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element.</span></span> <span data-ttu-id="bffce-137">Атрибут `resid` элемента **Описание может** быть не более 32 символов и устанавливается на `id` значение атрибута `String` элемента, который содержит текст.</span><span class="sxs-lookup"><span data-stu-id="bffce-137">The `resid` attribute of the **Description** element can be no more than 32 characters and is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="bffce-138">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="bffce-138">**Requirements**</span></span>  |  <span data-ttu-id="bffce-139">Нет</span><span class="sxs-lookup"><span data-stu-id="bffce-139">No</span></span>   |  <span data-ttu-id="bffce-p105">Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="bffce-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="bffce-142">Hosts</span><span class="sxs-lookup"><span data-stu-id="bffce-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="bffce-143">Да</span><span class="sxs-lookup"><span data-stu-id="bffce-143">Yes</span></span>  |  <span data-ttu-id="bffce-144">Определяет набор Office приложений.</span><span class="sxs-lookup"><span data-stu-id="bffce-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="bffce-145">Элемент «Хосты ребенка» перекрывает элемент «Хозяева» в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="bffce-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="bffce-146">Resources</span><span class="sxs-lookup"><span data-stu-id="bffce-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="bffce-147">Да</span><span class="sxs-lookup"><span data-stu-id="bffce-147">Yes</span></span>  | <span data-ttu-id="bffce-148">Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.</span><span class="sxs-lookup"><span data-stu-id="bffce-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="bffce-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="bffce-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="bffce-150">Нет</span><span class="sxs-lookup"><span data-stu-id="bffce-150">No</span></span>  | <span data-ttu-id="bffce-151">Определяет родные (COM/XLL) дополнения, эквивалентные веб-надстройки.</span><span class="sxs-lookup"><span data-stu-id="bffce-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="bffce-152">Веб-надстройок не активируется, если установлена эквивалентная пристройная система.</span><span class="sxs-lookup"><span data-stu-id="bffce-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="bffce-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="bffce-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="bffce-154">Нет</span><span class="sxs-lookup"><span data-stu-id="bffce-154">No</span></span>  | <span data-ttu-id="bffce-p108">Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="bffce-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="bffce-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="bffce-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="bffce-158">Нет</span><span class="sxs-lookup"><span data-stu-id="bffce-158">No</span></span>  | <span data-ttu-id="bffce-159">Уточняется подробная информация о регистрации надстройки с защищенными эмитентами токенов, такими как Azure Active Directory V2.0.</span><span class="sxs-lookup"><span data-stu-id="bffce-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="bffce-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="bffce-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="bffce-161">Нет</span><span class="sxs-lookup"><span data-stu-id="bffce-161">No</span></span>  |  <span data-ttu-id="bffce-162">Определяет набор расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="bffce-162">Specifies a collection of extended permissions.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="bffce-163">Пример VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="bffce-163">VersionOverrides example</span></span>

<span data-ttu-id="bffce-164">Ниже приводится пример типичного `<VersionOverrides>` элемента, включая некоторые элементы ребенка, которые не требуются, но обычно используются.</span><span class="sxs-lookup"><span data-stu-id="bffce-164">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="bffce-165">Реализация нескольких версий</span><span class="sxs-lookup"><span data-stu-id="bffce-165">Implementing multiple versions</span></span>

<span data-ttu-id="bffce-p109">В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.</span><span class="sxs-lookup"><span data-stu-id="bffce-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="bffce-168">Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии.</span><span class="sxs-lookup"><span data-stu-id="bffce-168">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="bffce-169">Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.</span><span class="sxs-lookup"><span data-stu-id="bffce-169">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="bffce-170">Чтобы реализовать схему VersionOverrides версий 1.0 и 1.1, манифест должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="bffce-170">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
