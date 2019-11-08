---
title: Элемент Host в файле манифеста
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 824cc6ae51eb9db713a0a9a768e3ec48e3271e95
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066279"
---
# <a name="host-element"></a><span data-ttu-id="16777-102">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="16777-102">Host element</span></span>

<span data-ttu-id="16777-103">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="16777-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="16777-104">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="16777-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="16777-105">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="16777-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="16777-106">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="16777-106">Basic manifest</span></span>

<span data-ttu-id="16777-107">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="16777-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="16777-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="16777-108">Attributes</span></span>

| <span data-ttu-id="16777-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="16777-109">Attribute</span></span>     | <span data-ttu-id="16777-110">Тип</span><span class="sxs-lookup"><span data-stu-id="16777-110">Type</span></span>   | <span data-ttu-id="16777-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="16777-111">Required</span></span> | <span data-ttu-id="16777-112">Описание</span><span class="sxs-lookup"><span data-stu-id="16777-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="16777-113">Name</span><span class="sxs-lookup"><span data-stu-id="16777-113">Name</span></span>](#name) | <span data-ttu-id="16777-114">string</span><span class="sxs-lookup"><span data-stu-id="16777-114">string</span></span> | <span data-ttu-id="16777-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="16777-115">required</span></span> | <span data-ttu-id="16777-116">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="16777-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="16777-117">Имя</span><span class="sxs-lookup"><span data-stu-id="16777-117">Name</span></span>

<span data-ttu-id="16777-118">Определяет тип ведущего приложения, для которого предназначена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="16777-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="16777-119">Значение должно быть одним из следующих.</span><span class="sxs-lookup"><span data-stu-id="16777-119">The value must be one of the following.</span></span>

- <span data-ttu-id="16777-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="16777-120">`Document` (Word)</span></span>
- <span data-ttu-id="16777-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="16777-121">`Database` (Access)</span></span>
- <span data-ttu-id="16777-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="16777-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="16777-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="16777-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="16777-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="16777-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="16777-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="16777-125">`Project` (Project)</span></span>
- <span data-ttu-id="16777-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="16777-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="16777-127">Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint.</span><span class="sxs-lookup"><span data-stu-id="16777-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="16777-128">В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.</span><span class="sxs-lookup"><span data-stu-id="16777-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="16777-129">Пример</span><span class="sxs-lookup"><span data-stu-id="16777-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="16777-130">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="16777-130">VersionOverrides node</span></span>

<span data-ttu-id="16777-131">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="16777-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="16777-132">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="16777-132">Attributes</span></span>

|  <span data-ttu-id="16777-133">Атрибут</span><span class="sxs-lookup"><span data-stu-id="16777-133">Attribute</span></span>  |  <span data-ttu-id="16777-134">Обязательный</span><span class="sxs-lookup"><span data-stu-id="16777-134">Required</span></span>  |  <span data-ttu-id="16777-135">Описание</span><span class="sxs-lookup"><span data-stu-id="16777-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="16777-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="16777-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="16777-137">Да</span><span class="sxs-lookup"><span data-stu-id="16777-137">Yes</span></span>  | <span data-ttu-id="16777-138">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="16777-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="16777-139">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="16777-139">Child elements</span></span>

|  <span data-ttu-id="16777-140">Элемент</span><span class="sxs-lookup"><span data-stu-id="16777-140">Element</span></span> |  <span data-ttu-id="16777-141">Обязательный</span><span class="sxs-lookup"><span data-stu-id="16777-141">Required</span></span>  |  <span data-ttu-id="16777-142">Описание</span><span class="sxs-lookup"><span data-stu-id="16777-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="16777-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="16777-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="16777-144">Да</span><span class="sxs-lookup"><span data-stu-id="16777-144">Yes</span></span>   |  <span data-ttu-id="16777-145">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="16777-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="16777-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="16777-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="16777-147">Нет</span><span class="sxs-lookup"><span data-stu-id="16777-147">No</span></span>   |  <span data-ttu-id="16777-148">Определяет параметры для мобильного конструктивного параметра.</span><span class="sxs-lookup"><span data-stu-id="16777-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="16777-149">**Примечание:** Этот элемент поддерживается только в Outlook на iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="16777-149">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="16777-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="16777-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="16777-151">Нет</span><span class="sxs-lookup"><span data-stu-id="16777-151">No</span></span>   |  <span data-ttu-id="16777-152">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="16777-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="16777-153">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="16777-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="16777-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="16777-154">xsi:type</span></span>

<span data-ttu-id="16777-155">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="16777-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="16777-156">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="16777-156">The value must be one of the following:</span></span>

- <span data-ttu-id="16777-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="16777-157">`Document` (Word)</span></span>
- <span data-ttu-id="16777-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="16777-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="16777-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="16777-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="16777-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="16777-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="16777-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="16777-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="16777-162">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="16777-162">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
