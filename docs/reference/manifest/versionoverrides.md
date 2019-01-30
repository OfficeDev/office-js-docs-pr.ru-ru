---
title: Элемент VersionOverrides в файле манифеста
description: ''
ms.date: 01/29/2019
localization_priority: Normal
ms.openlocfilehash: 897c2203ef6ae84911b7f269ee8a2c88aec36bd0
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635911"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="73bac-102">Элемент VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="73bac-102">VersionOverrides element</span></span>

<span data-ttu-id="73bac-p101">Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.</span><span class="sxs-lookup"><span data-stu-id="73bac-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="73bac-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="73bac-106">Attributes</span></span>

|  <span data-ttu-id="73bac-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="73bac-107">Attribute</span></span>  |  <span data-ttu-id="73bac-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="73bac-108">Required</span></span>  |  <span data-ttu-id="73bac-109">Описание</span><span class="sxs-lookup"><span data-stu-id="73bac-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="73bac-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="73bac-110">**xmlns**</span></span>       |  <span data-ttu-id="73bac-111">Да</span><span class="sxs-lookup"><span data-stu-id="73bac-111">Yes</span></span>  |  <span data-ttu-id="73bac-112">Расположение схемы (`http://schemas.microsoft.com/office/mailappversionoverrides`, когда `xsi:type` — `VersionOverridesV1_0`, и `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, когда `xsi:type` — `VersionOverridesV1_1`).</span><span class="sxs-lookup"><span data-stu-id="73bac-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="73bac-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="73bac-113">**xsi:type**</span></span>  |  <span data-ttu-id="73bac-114">Да</span><span class="sxs-lookup"><span data-stu-id="73bac-114">Yes</span></span>  | <span data-ttu-id="73bac-p102">Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="73bac-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="73bac-117">В настоящее время только Outlook 2016 или более поздняя версия поддерживает схемы версии 1.1 VersionOverrides и `VersionOverridesV1_1` типа.</span><span class="sxs-lookup"><span data-stu-id="73bac-117">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="73bac-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="73bac-118">Child elements</span></span>

|  <span data-ttu-id="73bac-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="73bac-119">Element</span></span> |  <span data-ttu-id="73bac-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="73bac-120">Required</span></span>  |  <span data-ttu-id="73bac-121">Описание</span><span class="sxs-lookup"><span data-stu-id="73bac-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="73bac-122">**Описание**</span><span class="sxs-lookup"><span data-stu-id="73bac-122">**Description**</span></span>    |  <span data-ttu-id="73bac-123">НЕТ</span><span class="sxs-lookup"><span data-stu-id="73bac-123">No</span></span>   |  <span data-ttu-id="73bac-p103">Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](./resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.</span><span class="sxs-lookup"><span data-stu-id="73bac-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="73bac-128">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="73bac-128">**Requirements**</span></span>  |  <span data-ttu-id="73bac-129">Нет</span><span class="sxs-lookup"><span data-stu-id="73bac-129">No</span></span>   |  <span data-ttu-id="73bac-p104">Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="73bac-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="73bac-132">Hosts</span><span class="sxs-lookup"><span data-stu-id="73bac-132">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="73bac-133">Да</span><span class="sxs-lookup"><span data-stu-id="73bac-133">Yes</span></span>  |  <span data-ttu-id="73bac-p105">Задает набор узлов Office. Дочерний элемент Hosts переопределяет элемент Hosts в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="73bac-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="73bac-136">Resources</span><span class="sxs-lookup"><span data-stu-id="73bac-136">Resources</span></span>](./resources.md)    |  <span data-ttu-id="73bac-137">Да</span><span class="sxs-lookup"><span data-stu-id="73bac-137">Yes</span></span>  | <span data-ttu-id="73bac-138">Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.</span><span class="sxs-lookup"><span data-stu-id="73bac-138">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="73bac-139">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="73bac-139">**VersionOverrides**</span></span>    |  <span data-ttu-id="73bac-140">Нет</span><span class="sxs-lookup"><span data-stu-id="73bac-140">No</span></span>  | <span data-ttu-id="73bac-p106">Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="73bac-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="73bac-143">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="73bac-143">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="73bac-144">Нет</span><span class="sxs-lookup"><span data-stu-id="73bac-144">No</span></span>  | <span data-ttu-id="73bac-145">Указывает сведения о связанном с надстройкой веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="73bac-145">Specifies details about the add-in's associated Web application.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="73bac-146">Пример VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="73bac-146">VersionOverrides example</span></span>

<span data-ttu-id="73bac-147">Ниже приведен пример типичного `<VersionOverrides>` элемент, включая некоторые дочерние элементы, которые не являются обязательными, но обычно используются.</span><span class="sxs-lookup"><span data-stu-id="73bac-147">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp>
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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="73bac-148">Реализация нескольких версий</span><span class="sxs-lookup"><span data-stu-id="73bac-148">Implementing multiple versions</span></span>

<span data-ttu-id="73bac-p107">В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.</span><span class="sxs-lookup"><span data-stu-id="73bac-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="73bac-151">Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии.</span><span class="sxs-lookup"><span data-stu-id="73bac-151">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="73bac-152">Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.</span><span class="sxs-lookup"><span data-stu-id="73bac-152">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="73bac-153">Чтобы реализовать схему VersionOverrides версий 1.0 и 1.1, манифест должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="73bac-153">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
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
