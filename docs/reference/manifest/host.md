---
title: Элемент Host в файле манифеста
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: e7b557034f70b03ed57598b7ffb9f43878db7392
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454897"
---
# <a name="host-element"></a><span data-ttu-id="1b681-102">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="1b681-102">Host element</span></span>

<span data-ttu-id="1b681-103">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="1b681-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="1b681-104">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="1b681-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="1b681-105">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="1b681-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="1b681-106">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="1b681-106">Basic manifest</span></span>

<span data-ttu-id="1b681-107">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="1b681-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="1b681-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1b681-108">Attributes</span></span>

| <span data-ttu-id="1b681-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="1b681-109">Attribute</span></span>     | <span data-ttu-id="1b681-110">Тип</span><span class="sxs-lookup"><span data-stu-id="1b681-110">Type</span></span>   | <span data-ttu-id="1b681-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1b681-111">Required</span></span> | <span data-ttu-id="1b681-112">Описание</span><span class="sxs-lookup"><span data-stu-id="1b681-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="1b681-113">Name</span><span class="sxs-lookup"><span data-stu-id="1b681-113">Name</span></span>](#name) | <span data-ttu-id="1b681-114">string</span><span class="sxs-lookup"><span data-stu-id="1b681-114">string</span></span> | <span data-ttu-id="1b681-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1b681-115">required</span></span> | <span data-ttu-id="1b681-116">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="1b681-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="1b681-117">Имя</span><span class="sxs-lookup"><span data-stu-id="1b681-117">Name</span></span>

<span data-ttu-id="1b681-118">Определяет тип ведущего приложения, для которого предназначена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="1b681-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="1b681-119">Значение должно быть одним из следующих.</span><span class="sxs-lookup"><span data-stu-id="1b681-119">The value must be one of the following.</span></span>

- <span data-ttu-id="1b681-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="1b681-120">`Document` (Word)</span></span>
- <span data-ttu-id="1b681-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="1b681-121">`Database` (Access)</span></span>
- <span data-ttu-id="1b681-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="1b681-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="1b681-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="1b681-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="1b681-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="1b681-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="1b681-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="1b681-125">`Project` (Project)</span></span>
- <span data-ttu-id="1b681-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="1b681-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1b681-127">Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint.</span><span class="sxs-lookup"><span data-stu-id="1b681-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="1b681-128">В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.</span><span class="sxs-lookup"><span data-stu-id="1b681-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="1b681-129">Пример</span><span class="sxs-lookup"><span data-stu-id="1b681-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="1b681-130">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="1b681-130">VersionOverrides node</span></span>

<span data-ttu-id="1b681-131">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="1b681-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="1b681-132">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1b681-132">Attributes</span></span>

|  <span data-ttu-id="1b681-133">Атрибут</span><span class="sxs-lookup"><span data-stu-id="1b681-133">Attribute</span></span>  |  <span data-ttu-id="1b681-134">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1b681-134">Required</span></span>  |  <span data-ttu-id="1b681-135">Описание</span><span class="sxs-lookup"><span data-stu-id="1b681-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1b681-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1b681-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="1b681-137">Да</span><span class="sxs-lookup"><span data-stu-id="1b681-137">Yes</span></span>  | <span data-ttu-id="1b681-138">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="1b681-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="1b681-139">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="1b681-139">Child elements</span></span>

|  <span data-ttu-id="1b681-140">Элемент</span><span class="sxs-lookup"><span data-stu-id="1b681-140">Element</span></span> |  <span data-ttu-id="1b681-141">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1b681-141">Required</span></span>  |  <span data-ttu-id="1b681-142">Описание</span><span class="sxs-lookup"><span data-stu-id="1b681-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1b681-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="1b681-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="1b681-144">Да</span><span class="sxs-lookup"><span data-stu-id="1b681-144">Yes</span></span>   |  <span data-ttu-id="1b681-145">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="1b681-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="1b681-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="1b681-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="1b681-147">Нет</span><span class="sxs-lookup"><span data-stu-id="1b681-147">No</span></span>   |  <span data-ttu-id="1b681-148">Определяет параметры для мобильного конструктивного параметра.</span><span class="sxs-lookup"><span data-stu-id="1b681-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="1b681-149">**Примечание:** Этот элемент поддерживается только в Outlook в iOS.</span><span class="sxs-lookup"><span data-stu-id="1b681-149">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="1b681-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="1b681-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="1b681-151">Нет</span><span class="sxs-lookup"><span data-stu-id="1b681-151">No</span></span>   |  <span data-ttu-id="1b681-152">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="1b681-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="1b681-153">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="1b681-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="1b681-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1b681-154">xsi:type</span></span>

<span data-ttu-id="1b681-155">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="1b681-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="1b681-156">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="1b681-156">The value must be one of the following:</span></span>

- <span data-ttu-id="1b681-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="1b681-157">`Document` (Word)</span></span>
- <span data-ttu-id="1b681-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="1b681-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="1b681-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="1b681-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="1b681-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="1b681-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="1b681-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="1b681-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="1b681-162">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="1b681-162">Host example</span></span> 

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
