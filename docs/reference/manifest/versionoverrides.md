---
title: Элемент VersionOverrides в файле манифеста
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: ce65cdced1b3cf885cee09732c2cda0081a53cfc
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477882"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="98ec4-102">Элемент VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="98ec4-102">VersionOverrides element</span></span>

<span data-ttu-id="98ec4-p101">Корневой элемент, который содержит сведения о командах надстройки. Элемент манифеста **VersionOverrides** является дочерним для элемента [OfficeApp](./officeapp.md). Этот элемент поддерживается в схеме манифестов версий 1.1 и выше, но определяется в схеме VersionOverrides версии 1.0 или 1.1.</span><span class="sxs-lookup"><span data-stu-id="98ec4-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="98ec4-106">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="98ec4-106">Attributes</span></span>

|  <span data-ttu-id="98ec4-107">Атрибут</span><span class="sxs-lookup"><span data-stu-id="98ec4-107">Attribute</span></span>  |  <span data-ttu-id="98ec4-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="98ec4-108">Required</span></span>  |  <span data-ttu-id="98ec4-109">Описание</span><span class="sxs-lookup"><span data-stu-id="98ec4-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="98ec4-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="98ec4-110">**xmlns**</span></span>       |  <span data-ttu-id="98ec4-111">Да</span><span class="sxs-lookup"><span data-stu-id="98ec4-111">Yes</span></span>  |  <span data-ttu-id="98ec4-112">Расположение схемы (`http://schemas.microsoft.com/office/mailappversionoverrides`, когда `xsi:type` — `VersionOverridesV1_0`, и `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, когда `xsi:type` — `VersionOverridesV1_1`).</span><span class="sxs-lookup"><span data-stu-id="98ec4-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="98ec4-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="98ec4-113">**xsi:type**</span></span>  |  <span data-ttu-id="98ec4-114">Да</span><span class="sxs-lookup"><span data-stu-id="98ec4-114">Yes</span></span>  | <span data-ttu-id="98ec4-p102">Версия схемы. В настоящее время допускаются только значения `VersionOverridesV1_0` и `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="98ec4-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="98ec4-117">В настоящее время только Outlook 2016 или более поздней версии поддерживает схему VersionOverrides `VersionOverridesV1_1` 1.1 и тип.</span><span class="sxs-lookup"><span data-stu-id="98ec4-117">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="98ec4-118">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="98ec4-118">Child elements</span></span>

|  <span data-ttu-id="98ec4-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="98ec4-119">Element</span></span> |  <span data-ttu-id="98ec4-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="98ec4-120">Required</span></span>  |  <span data-ttu-id="98ec4-121">Описание</span><span class="sxs-lookup"><span data-stu-id="98ec4-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="98ec4-122">**Описание**</span><span class="sxs-lookup"><span data-stu-id="98ec4-122">**Description**</span></span>    |  <span data-ttu-id="98ec4-123">НЕТ</span><span class="sxs-lookup"><span data-stu-id="98ec4-123">No</span></span>   |  <span data-ttu-id="98ec4-p103">Описывает надстройку. Переопределяет элемент `Description` в любой родительской части манифеста. Текст описания содержится в дочернем элементе **LongString**, включенном в элемент [Resources](./resources.md). Для атрибута `resid` элемента **Description** задано значение атрибута `id` элемента `String`, который содержит текст.</span><span class="sxs-lookup"><span data-stu-id="98ec4-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
| <span data-ttu-id="98ec4-128">**екуивалентаддинс**</span><span class="sxs-lookup"><span data-stu-id="98ec4-128">**EquivalentAddins**</span></span> | <span data-ttu-id="98ec4-129">Нет</span><span class="sxs-lookup"><span data-stu-id="98ec4-129">No</span></span> | <span data-ttu-id="98ec4-130">Задает обратную совместимость с эквивалентной надстройкой COM, XLL или и тем, и другими.</span><span class="sxs-lookup"><span data-stu-id="98ec4-130">Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.</span></span> |
|  <span data-ttu-id="98ec4-131">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="98ec4-131">**Requirements**</span></span>  |  <span data-ttu-id="98ec4-132">Нет</span><span class="sxs-lookup"><span data-stu-id="98ec4-132">No</span></span>   |  <span data-ttu-id="98ec4-p104">Задает минимальные набор требований и версию библиотеки Office.js, необходимые надстройке. Переопределяет элемент `Requirements` в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="98ec4-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="98ec4-135">Hosts</span><span class="sxs-lookup"><span data-stu-id="98ec4-135">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="98ec4-136">Да</span><span class="sxs-lookup"><span data-stu-id="98ec4-136">Yes</span></span>  |  <span data-ttu-id="98ec4-p105">Задает набор узлов Office. Дочерний элемент Hosts переопределяет элемент Hosts в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="98ec4-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="98ec4-139">Resources</span><span class="sxs-lookup"><span data-stu-id="98ec4-139">Resources</span></span>](./resources.md)    |  <span data-ttu-id="98ec4-140">Да</span><span class="sxs-lookup"><span data-stu-id="98ec4-140">Yes</span></span>  | <span data-ttu-id="98ec4-141">Определяет коллекцию ресурсов (строк, URL-адресов и изображений), на которые ссылаются другие элементы манифеста.</span><span class="sxs-lookup"><span data-stu-id="98ec4-141">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="98ec4-142">екуивалентаддинс</span><span class="sxs-lookup"><span data-stu-id="98ec4-142">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="98ec4-143">Нет</span><span class="sxs-lookup"><span data-stu-id="98ec4-143">No</span></span>  | <span data-ttu-id="98ec4-144">Задает встроенные надстройки (COM/XLL), эквивалентные веб-надстройке.</span><span class="sxs-lookup"><span data-stu-id="98ec4-144">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="98ec4-145">Веб-надстройка не активируется, если установлена эквивалентная собственная встроенная надстройка.</span><span class="sxs-lookup"><span data-stu-id="98ec4-145">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="98ec4-146">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="98ec4-146">**VersionOverrides**</span></span>    |  <span data-ttu-id="98ec4-147">Нет</span><span class="sxs-lookup"><span data-stu-id="98ec4-147">No</span></span>  | <span data-ttu-id="98ec4-p107">Определяет команды надстроек в новой версии схемы. Подробные сведения см. в разделе [Реализация нескольких версий](#implementing-multiple-versions).</span><span class="sxs-lookup"><span data-stu-id="98ec4-p107">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="98ec4-150">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="98ec4-150">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="98ec4-151">Нет</span><span class="sxs-lookup"><span data-stu-id="98ec4-151">No</span></span>  | <span data-ttu-id="98ec4-152">Задает сведения о регистрации надстройки с помощью надежных поставщиков маркеров, таких как Azure Active Directory 2.0.</span><span class="sxs-lookup"><span data-stu-id="98ec4-152">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="98ec4-153">Пример VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="98ec4-153">VersionOverrides example</span></span>

<span data-ttu-id="98ec4-154">Ниже приведен пример типичного `<VersionOverrides>` элемента, в том числе некоторые необязательные дочерние элементы, которые обычно используются.</span><span class="sxs-lookup"><span data-stu-id="98ec4-154">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="98ec4-155">Реализация нескольких версий</span><span class="sxs-lookup"><span data-stu-id="98ec4-155">Implementing multiple versions</span></span>

<span data-ttu-id="98ec4-p108">В манифесте может быть реализовано несколько версий элемента `VersionOverrides`, которые поддерживают различные версии схемы VersionOverrides. Это можно сделать, чтобы поддерживать новые функции в новой схеме, по-прежнему поддерживая старые клиенты.</span><span class="sxs-lookup"><span data-stu-id="98ec4-p108">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="98ec4-158">Чтобы реализовать несколько версий, элемент `VersionOverrides` для новой версии должен зависеть от элемента `VersionOverrides` для старой версии.</span><span class="sxs-lookup"><span data-stu-id="98ec4-158">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="98ec4-159">Дочерний элемент `VersionOverrides` не наследует значения от родительского объекта.</span><span class="sxs-lookup"><span data-stu-id="98ec4-159">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="98ec4-160">Чтобы реализовать схему VersionOverrides версий 1.0 и 1.1, манифест должен выглядеть следующим образом:</span><span class="sxs-lookup"><span data-stu-id="98ec4-160">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
