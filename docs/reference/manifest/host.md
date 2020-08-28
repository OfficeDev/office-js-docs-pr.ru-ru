---
title: Элемент Host в файле манифеста
description: Определяет тип приложения Office, в котором следует активировать надстройку.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5b6c6e6b5471b4117c28cf92e11eb0a99b512a97
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292288"
---
# <a name="host-element"></a><span data-ttu-id="ebf3f-103">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="ebf3f-103">Host element</span></span>

<span data-ttu-id="ebf3f-104">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ebf3f-105">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="ebf3f-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="ebf3f-106">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="ebf3f-107">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="ebf3f-107">Basic manifest</span></span>

<span data-ttu-id="ebf3f-108">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="ebf3f-109">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ebf3f-109">Attributes</span></span>

| <span data-ttu-id="ebf3f-110">Атрибут</span><span class="sxs-lookup"><span data-stu-id="ebf3f-110">Attribute</span></span>     | <span data-ttu-id="ebf3f-111">Тип</span><span class="sxs-lookup"><span data-stu-id="ebf3f-111">Type</span></span>   | <span data-ttu-id="ebf3f-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ebf3f-112">Required</span></span> | <span data-ttu-id="ebf3f-113">Описание</span><span class="sxs-lookup"><span data-stu-id="ebf3f-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="ebf3f-114">Name</span><span class="sxs-lookup"><span data-stu-id="ebf3f-114">Name</span></span>](#name) | <span data-ttu-id="ebf3f-115">string</span><span class="sxs-lookup"><span data-stu-id="ebf3f-115">string</span></span> | <span data-ttu-id="ebf3f-116">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ebf3f-116">required</span></span> | <span data-ttu-id="ebf3f-117">Имя типа клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="ebf3f-118">Имя</span><span class="sxs-lookup"><span data-stu-id="ebf3f-118">Name</span></span>

<span data-ttu-id="ebf3f-119">Определяет тип ведущего приложения, для которого предназначена эта надстройка.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="ebf3f-120">Значение должно быть одним из следующих.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-120">The value must be one of the following.</span></span>

- <span data-ttu-id="ebf3f-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-121">`Document` (Word)</span></span>
- <span data-ttu-id="ebf3f-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-122">`Database` (Access)</span></span>
- <span data-ttu-id="ebf3f-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="ebf3f-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ebf3f-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ebf3f-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-126">`Project` (Project)</span></span>
- <span data-ttu-id="ebf3f-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ebf3f-128">Больше не рекомендуется создавать и использовать веб-приложения и базы данных Access в SharePoint.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="ebf3f-129">В качестве альтернативы рекомендуем использовать [Microsoft PowerApps](https://powerapps.microsoft.com/) для создания бизнес-решений для Интернета и мобильных устройств без написания кода.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="ebf3f-130">Пример</span><span class="sxs-lookup"><span data-stu-id="ebf3f-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="ebf3f-131">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="ebf3f-131">VersionOverrides node</span></span>

<span data-ttu-id="ebf3f-132">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="ebf3f-133">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ebf3f-133">Attributes</span></span>

|  <span data-ttu-id="ebf3f-134">Атрибут</span><span class="sxs-lookup"><span data-stu-id="ebf3f-134">Attribute</span></span>  |  <span data-ttu-id="ebf3f-135">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ebf3f-135">Required</span></span>  |  <span data-ttu-id="ebf3f-136">Описание</span><span class="sxs-lookup"><span data-stu-id="ebf3f-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ebf3f-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ebf3f-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="ebf3f-138">Да</span><span class="sxs-lookup"><span data-stu-id="ebf3f-138">Yes</span></span>  | <span data-ttu-id="ebf3f-139">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="ebf3f-140">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ebf3f-140">Child elements</span></span>

|  <span data-ttu-id="ebf3f-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="ebf3f-141">Element</span></span> |  <span data-ttu-id="ebf3f-142">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ebf3f-142">Required</span></span>  |  <span data-ttu-id="ebf3f-143">Описание</span><span class="sxs-lookup"><span data-stu-id="ebf3f-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ebf3f-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ebf3f-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="ebf3f-145">Да</span><span class="sxs-lookup"><span data-stu-id="ebf3f-145">Yes</span></span>   |  <span data-ttu-id="ebf3f-146">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="ebf3f-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ebf3f-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="ebf3f-148">Нет</span><span class="sxs-lookup"><span data-stu-id="ebf3f-148">No</span></span>   |  <span data-ttu-id="ebf3f-149">Определяет параметры для мобильного конструктивного параметра.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="ebf3f-150">**Примечание:** Этот элемент поддерживается только в Outlook на iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="ebf3f-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ebf3f-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="ebf3f-152">Нет</span><span class="sxs-lookup"><span data-stu-id="ebf3f-152">No</span></span>   |  <span data-ttu-id="ebf3f-153">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="ebf3f-154">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="ebf3f-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ebf3f-155">xsi:type</span></span>

<span data-ttu-id="ebf3f-156">Управляет приложением Office (Word, Excel, PowerPoint, Outlook, OneNote), к которому применяются вложенные параметры.</span><span class="sxs-lookup"><span data-stu-id="ebf3f-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="ebf3f-157">Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="ebf3f-157">The value must be one of the following:</span></span>

- <span data-ttu-id="ebf3f-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-158">`Document` (Word)</span></span>
- <span data-ttu-id="ebf3f-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="ebf3f-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ebf3f-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ebf3f-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ebf3f-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="ebf3f-163">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="ebf3f-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
