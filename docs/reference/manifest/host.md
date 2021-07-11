---
title: Элемент Host в файле манифеста
description: Определяет тип приложения Office, в котором следует активировать надстройку.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 45d4ed42946038699be235ff3912c071a92ff226
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348330"
---
# <a name="host-element"></a><span data-ttu-id="a3931-103">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="a3931-103">Host element</span></span>

<span data-ttu-id="a3931-104">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="a3931-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a3931-105">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="a3931-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="a3931-106">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="a3931-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="a3931-107">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="a3931-107">Basic manifest</span></span>

<span data-ttu-id="a3931-108">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="a3931-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="a3931-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a3931-109">Attributes</span></span>

| <span data-ttu-id="a3931-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="a3931-110">Attribute</span></span>     | <span data-ttu-id="a3931-111">Тип</span><span class="sxs-lookup"><span data-stu-id="a3931-111">Type</span></span>   | <span data-ttu-id="a3931-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a3931-112">Required</span></span> | <span data-ttu-id="a3931-113">Описание</span><span class="sxs-lookup"><span data-stu-id="a3931-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="a3931-114">Name</span><span class="sxs-lookup"><span data-stu-id="a3931-114">Name</span></span>](#name) | <span data-ttu-id="a3931-115">string</span><span class="sxs-lookup"><span data-stu-id="a3931-115">string</span></span> | <span data-ttu-id="a3931-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a3931-116">required</span></span> | <span data-ttu-id="a3931-117">Имя типа Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="a3931-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="a3931-118">Имя</span><span class="sxs-lookup"><span data-stu-id="a3931-118">Name</span></span>

<span data-ttu-id="a3931-p102">Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="a3931-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="a3931-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="a3931-121">`Document` (Word)</span></span>
- <span data-ttu-id="a3931-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="a3931-122">`Database` (Access)</span></span>
- <span data-ttu-id="a3931-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="a3931-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="a3931-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="a3931-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="a3931-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="a3931-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="a3931-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="a3931-126">`Project` (Project)</span></span>
- <span data-ttu-id="a3931-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="a3931-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a3931-128">Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint.</span><span class="sxs-lookup"><span data-stu-id="a3931-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="a3931-129">В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.</span><span class="sxs-lookup"><span data-stu-id="a3931-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="a3931-130">Пример</span><span class="sxs-lookup"><span data-stu-id="a3931-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="a3931-131">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="a3931-131">VersionOverrides node</span></span>

<span data-ttu-id="a3931-132">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="a3931-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="a3931-133">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a3931-133">Attributes</span></span>

|  <span data-ttu-id="a3931-134">Атрибут</span><span class="sxs-lookup"><span data-stu-id="a3931-134">Attribute</span></span>  |  <span data-ttu-id="a3931-135">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a3931-135">Required</span></span>  |  <span data-ttu-id="a3931-136">Описание</span><span class="sxs-lookup"><span data-stu-id="a3931-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a3931-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a3931-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="a3931-138">Да</span><span class="sxs-lookup"><span data-stu-id="a3931-138">Yes</span></span>  | <span data-ttu-id="a3931-139">Описывает приложение Office, в котором применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="a3931-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="a3931-140">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="a3931-140">Child elements</span></span>

|  <span data-ttu-id="a3931-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="a3931-141">Element</span></span> |  <span data-ttu-id="a3931-142">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a3931-142">Required</span></span>  |  <span data-ttu-id="a3931-143">Описание</span><span class="sxs-lookup"><span data-stu-id="a3931-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a3931-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="a3931-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="a3931-145">Да</span><span class="sxs-lookup"><span data-stu-id="a3931-145">Yes</span></span>   |  <span data-ttu-id="a3931-146">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="a3931-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="a3931-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="a3931-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="a3931-148">Нет</span><span class="sxs-lookup"><span data-stu-id="a3931-148">No</span></span>   |  <span data-ttu-id="a3931-149">Определяет параметры мобильного форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="a3931-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="a3931-150">**Примечание:** Этот элемент поддерживается только в Outlook iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a3931-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="a3931-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="a3931-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="a3931-152">Нет</span><span class="sxs-lookup"><span data-stu-id="a3931-152">No</span></span>   |  <span data-ttu-id="a3931-153">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="a3931-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="a3931-154">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="a3931-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="a3931-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a3931-155">xsi:type</span></span>

<span data-ttu-id="a3931-156">Элементы управления Office приложения (Word, Excel, PowerPoint, Outlook, OneNote), где применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="a3931-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="a3931-157">Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="a3931-157">The value must be one of the following:</span></span>

- <span data-ttu-id="a3931-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="a3931-158">`Document` (Word)</span></span>
- <span data-ttu-id="a3931-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="a3931-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="a3931-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="a3931-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="a3931-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="a3931-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="a3931-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="a3931-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="a3931-163">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="a3931-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
