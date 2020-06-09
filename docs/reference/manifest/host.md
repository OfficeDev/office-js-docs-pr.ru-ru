---
title: Элемент Host в файле манифеста
description: Определяет тип приложения Office, в котором следует активировать надстройку.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5db9df97c4ba558d54756b983a26cb7b71e049d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611815"
---
# <a name="host-element"></a><span data-ttu-id="f9621-103">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="f9621-103">Host element</span></span>

<span data-ttu-id="f9621-104">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="f9621-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9621-105">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="f9621-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="f9621-106">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="f9621-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="f9621-107">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="f9621-107">Basic manifest</span></span>

<span data-ttu-id="f9621-108">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="f9621-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="f9621-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f9621-109">Attributes</span></span>

| <span data-ttu-id="f9621-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f9621-110">Attribute</span></span>     | <span data-ttu-id="f9621-111">Тип</span><span class="sxs-lookup"><span data-stu-id="f9621-111">Type</span></span>   | <span data-ttu-id="f9621-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f9621-112">Required</span></span> | <span data-ttu-id="f9621-113">Описание</span><span class="sxs-lookup"><span data-stu-id="f9621-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="f9621-114">Name</span><span class="sxs-lookup"><span data-stu-id="f9621-114">Name</span></span>](#name) | <span data-ttu-id="f9621-115">string</span><span class="sxs-lookup"><span data-stu-id="f9621-115">string</span></span> | <span data-ttu-id="f9621-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f9621-116">required</span></span> | <span data-ttu-id="f9621-117">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="f9621-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="f9621-118">Имя</span><span class="sxs-lookup"><span data-stu-id="f9621-118">Name</span></span>

<span data-ttu-id="f9621-119">Определяет тип ведущего приложения, для которого предназначена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="f9621-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="f9621-120">Значение должно быть одним из следующих.</span><span class="sxs-lookup"><span data-stu-id="f9621-120">The value must be one of the following.</span></span>

- <span data-ttu-id="f9621-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="f9621-121">`Document` (Word)</span></span>
- <span data-ttu-id="f9621-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="f9621-122">`Database` (Access)</span></span>
- <span data-ttu-id="f9621-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="f9621-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="f9621-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="f9621-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="f9621-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="f9621-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="f9621-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="f9621-126">`Project` (Project)</span></span>
- <span data-ttu-id="f9621-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="f9621-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f9621-128">Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint.</span><span class="sxs-lookup"><span data-stu-id="f9621-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="f9621-129">В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.</span><span class="sxs-lookup"><span data-stu-id="f9621-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="f9621-130">Пример</span><span class="sxs-lookup"><span data-stu-id="f9621-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="f9621-131">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="f9621-131">VersionOverrides node</span></span>

<span data-ttu-id="f9621-132">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="f9621-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="f9621-133">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f9621-133">Attributes</span></span>

|  <span data-ttu-id="f9621-134">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f9621-134">Attribute</span></span>  |  <span data-ttu-id="f9621-135">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f9621-135">Required</span></span>  |  <span data-ttu-id="f9621-136">Описание</span><span class="sxs-lookup"><span data-stu-id="f9621-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f9621-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f9621-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="f9621-138">Да</span><span class="sxs-lookup"><span data-stu-id="f9621-138">Yes</span></span>  | <span data-ttu-id="f9621-139">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="f9621-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="f9621-140">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="f9621-140">Child elements</span></span>

|  <span data-ttu-id="f9621-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="f9621-141">Element</span></span> |  <span data-ttu-id="f9621-142">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f9621-142">Required</span></span>  |  <span data-ttu-id="f9621-143">Описание</span><span class="sxs-lookup"><span data-stu-id="f9621-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f9621-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f9621-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="f9621-145">Да</span><span class="sxs-lookup"><span data-stu-id="f9621-145">Yes</span></span>   |  <span data-ttu-id="f9621-146">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="f9621-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="f9621-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="f9621-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="f9621-148">Нет</span><span class="sxs-lookup"><span data-stu-id="f9621-148">No</span></span>   |  <span data-ttu-id="f9621-149">Определяет параметры для мобильного конструктивного параметра.</span><span class="sxs-lookup"><span data-stu-id="f9621-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="f9621-150">**Примечание:** Этот элемент поддерживается только в Outlook на iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f9621-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="f9621-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="f9621-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="f9621-152">Нет</span><span class="sxs-lookup"><span data-stu-id="f9621-152">No</span></span>   |  <span data-ttu-id="f9621-153">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="f9621-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="f9621-154">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="f9621-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="f9621-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f9621-155">xsi:type</span></span>

<span data-ttu-id="f9621-156">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="f9621-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="f9621-157">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="f9621-157">The value must be one of the following:</span></span>

- <span data-ttu-id="f9621-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="f9621-158">`Document` (Word)</span></span>
- <span data-ttu-id="f9621-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="f9621-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="f9621-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="f9621-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="f9621-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="f9621-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="f9621-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="f9621-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="f9621-163">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="f9621-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
