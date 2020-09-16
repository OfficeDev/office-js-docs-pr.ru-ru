---
title: Элемент VersionOverrides в файле манифеста
description: Справочная документация по элементу VersionOverrides для файлов манифеста надстроек Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: a744772c01c57c41a9dc20ee0accea5f070c3ff3
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819828"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="93a38-103">Элемент VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="93a38-103">VersionOverrides element</span></span>

<span data-ttu-id="93a38-p101">Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.</span><span class="sxs-lookup"><span data-stu-id="93a38-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="93a38-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="93a38-107">Attributes</span></span>

|  <span data-ttu-id="93a38-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="93a38-108">Attribute</span></span>  |  <span data-ttu-id="93a38-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="93a38-109">Required</span></span>  |  <span data-ttu-id="93a38-110">Описание</span><span class="sxs-lookup"><span data-stu-id="93a38-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="93a38-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="93a38-111">**xmlns**</span></span>       |  <span data-ttu-id="93a38-112">Да</span><span class="sxs-lookup"><span data-stu-id="93a38-112">Yes</span></span>  |  <span data-ttu-id="93a38-113">Пространство имен схемы VersionOverrides.</span><span class="sxs-lookup"><span data-stu-id="93a38-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="93a38-114">Допустимые значения зависят от `<VersionOverrides>` значения **xsi: Type** этого элемента и значения **xsi: Type** родительского `<OfficeApp>` элемента.</span><span class="sxs-lookup"><span data-stu-id="93a38-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="93a38-115">Ниже приведены [значения пространств имен](#namespace-values) .</span><span class="sxs-lookup"><span data-stu-id="93a38-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="93a38-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="93a38-116">**xsi:type**</span></span>  |  <span data-ttu-id="93a38-117">Да</span><span class="sxs-lookup"><span data-stu-id="93a38-117">Yes</span></span>  | <span data-ttu-id="93a38-p103">Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="93a38-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="93a38-120">Значения пространств имен</span><span class="sxs-lookup"><span data-stu-id="93a38-120">Namespace values</span></span>

<span data-ttu-id="93a38-121">Ниже приведен список требуемого значения **xmlns** в зависимости от значения **xsi: Type** родительского `<OfficeApp>` элемента.</span><span class="sxs-lookup"><span data-stu-id="93a38-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="93a38-122">**TaskPaneApp** поддерживает только версию 1,0 VersionOverrides, а **xmlns** — значение `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="93a38-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="93a38-123">**ContentApp** поддерживает только версию 1,0 VersionOverrides, а **xmlns** — значение `http://schemas.microsoft.com/office/contentappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="93a38-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="93a38-124">**MailApp** поддерживает версии 1,0 и 1,1 для VersionOverrides, поэтому значение **xmlns** зависит от `<VersionOverrides>` значения **xsi: Type** этого элемента:</span><span class="sxs-lookup"><span data-stu-id="93a38-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="93a38-125">Если **xsi: Type** `VersionOverridesV1_0` , то **xmlns** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="93a38-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="93a38-126">Если **xsi: Type** `VersionOverridesV1_1` , то **xmlns** должен быть `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .</span><span class="sxs-lookup"><span data-stu-id="93a38-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="93a38-127">В настоящее время только Outlook 2016 или более поздней версии поддерживает схему VersionOverrides 1.1 и `VersionOverridesV1_1` тип.</span><span class="sxs-lookup"><span data-stu-id="93a38-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="93a38-128">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="93a38-128">Child elements</span></span>

|  <span data-ttu-id="93a38-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="93a38-129">Element</span></span> |  <span data-ttu-id="93a38-130">Обязательный</span><span class="sxs-lookup"><span data-stu-id="93a38-130">Required</span></span>  |  <span data-ttu-id="93a38-131">Описание</span><span class="sxs-lookup"><span data-stu-id="93a38-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="93a38-132">**Описание**</span><span class="sxs-lookup"><span data-stu-id="93a38-132">**Description**</span></span>    |  <span data-ttu-id="93a38-133">Нет</span><span class="sxs-lookup"><span data-stu-id="93a38-133">No</span></span>   |  <span data-ttu-id="93a38-p104">Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.</span><span class="sxs-lookup"><span data-stu-id="93a38-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="93a38-138">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="93a38-138">**Requirements**</span></span>  |  <span data-ttu-id="93a38-139">Нет</span><span class="sxs-lookup"><span data-stu-id="93a38-139">No</span></span>   |  <span data-ttu-id="93a38-p105">Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="93a38-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="93a38-142">Hosts</span><span class="sxs-lookup"><span data-stu-id="93a38-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="93a38-143">Да</span><span class="sxs-lookup"><span data-stu-id="93a38-143">Yes</span></span>  |  <span data-ttu-id="93a38-144">Задает коллекцию приложений Office.</span><span class="sxs-lookup"><span data-stu-id="93a38-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="93a38-145">Дочерний элемент hosts переопределяет элемент hosts в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="93a38-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="93a38-146">Resources</span><span class="sxs-lookup"><span data-stu-id="93a38-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="93a38-147">Да</span><span class="sxs-lookup"><span data-stu-id="93a38-147">Yes</span></span>  | <span data-ttu-id="93a38-148">Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.</span><span class="sxs-lookup"><span data-stu-id="93a38-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="93a38-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="93a38-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="93a38-150">Нет</span><span class="sxs-lookup"><span data-stu-id="93a38-150">No</span></span>  | <span data-ttu-id="93a38-151">Задает встроенные надстройки (COM/XLL), эквивалентные веб-надстройке.</span><span class="sxs-lookup"><span data-stu-id="93a38-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="93a38-152">Веб-надстройка не активируется, если установлена эквивалентная собственная встроенная надстройка.</span><span class="sxs-lookup"><span data-stu-id="93a38-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="93a38-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="93a38-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="93a38-154">Нет</span><span class="sxs-lookup"><span data-stu-id="93a38-154">No</span></span>  | <span data-ttu-id="93a38-p108">Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="93a38-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="93a38-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="93a38-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="93a38-158">Нет</span><span class="sxs-lookup"><span data-stu-id="93a38-158">No</span></span>  | <span data-ttu-id="93a38-159">Задает сведения о регистрации надстройки с помощью надежных поставщиков маркеров, таких как Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="93a38-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="93a38-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="93a38-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="93a38-161">Нет</span><span class="sxs-lookup"><span data-stu-id="93a38-161">No</span></span>  |  <span data-ttu-id="93a38-162">Задает коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="93a38-162">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="93a38-163">**Важно!** поскольку API [Office. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) в настоящее время находится в режиме предварительной версии, надстройки, использующие этот `ExtendedPermissions` элемент, не могут быть опубликованы в AppSource или развернуты с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="93a38-163">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="93a38-164">Пример VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="93a38-164">VersionOverrides example</span></span>

<span data-ttu-id="93a38-165">Ниже приведен пример типичного `<VersionOverrides>` элемента, в том числе некоторые необязательные дочерние элементы, которые обычно используются.</span><span class="sxs-lookup"><span data-stu-id="93a38-165">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="93a38-166">Реализация нескольких версий</span><span class="sxs-lookup"><span data-stu-id="93a38-166">Implementing multiple versions</span></span>

<span data-ttu-id="93a38-p109">В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.</span><span class="sxs-lookup"><span data-stu-id="93a38-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="93a38-169">Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии.</span><span class="sxs-lookup"><span data-stu-id="93a38-169">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="93a38-170">Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.</span><span class="sxs-lookup"><span data-stu-id="93a38-170">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="93a38-171">Чтобы реализовать схему VersionOverrides версий 1.0 и 1.1, манифест должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="93a38-171">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
