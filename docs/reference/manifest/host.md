---
title: Элемент Host в файле манифеста
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127571"
---
# <a name="host-element"></a><span data-ttu-id="8ca62-102">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="8ca62-102">Host element</span></span>

<span data-ttu-id="8ca62-103">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="8ca62-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="8ca62-104">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="8ca62-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="8ca62-105">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="8ca62-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="8ca62-106">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="8ca62-106">Basic manifest</span></span>

<span data-ttu-id="8ca62-107">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="8ca62-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="8ca62-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8ca62-108">Attributes</span></span>

| <span data-ttu-id="8ca62-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="8ca62-109">Attribute</span></span>     | <span data-ttu-id="8ca62-110">Тип</span><span class="sxs-lookup"><span data-stu-id="8ca62-110">Type</span></span>   | <span data-ttu-id="8ca62-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8ca62-111">Required</span></span> | <span data-ttu-id="8ca62-112">Описание</span><span class="sxs-lookup"><span data-stu-id="8ca62-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="8ca62-113">Name</span><span class="sxs-lookup"><span data-stu-id="8ca62-113">Name</span></span>](#name) | <span data-ttu-id="8ca62-114">string</span><span class="sxs-lookup"><span data-stu-id="8ca62-114">string</span></span> | <span data-ttu-id="8ca62-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8ca62-115">required</span></span> | <span data-ttu-id="8ca62-116">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="8ca62-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="8ca62-117">Имя</span><span class="sxs-lookup"><span data-stu-id="8ca62-117">Name</span></span>
<span data-ttu-id="8ca62-p102">Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="8ca62-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="8ca62-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="8ca62-120">`Document` (Word)</span></span>
- <span data-ttu-id="8ca62-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="8ca62-121">`Database` (Access)</span></span>
- <span data-ttu-id="8ca62-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="8ca62-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="8ca62-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="8ca62-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="8ca62-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="8ca62-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="8ca62-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="8ca62-125">`Project` (Project)</span></span>
- <span data-ttu-id="8ca62-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="8ca62-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="8ca62-127">Пример</span><span class="sxs-lookup"><span data-stu-id="8ca62-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="8ca62-128">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="8ca62-128">VersionOverrides node</span></span>
<span data-ttu-id="8ca62-129">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="8ca62-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="8ca62-130">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8ca62-130">Attributes</span></span>

|  <span data-ttu-id="8ca62-131">Атрибут</span><span class="sxs-lookup"><span data-stu-id="8ca62-131">Attribute</span></span>  |  <span data-ttu-id="8ca62-132">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8ca62-132">Required</span></span>  |  <span data-ttu-id="8ca62-133">Описание</span><span class="sxs-lookup"><span data-stu-id="8ca62-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8ca62-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8ca62-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="8ca62-135">Да</span><span class="sxs-lookup"><span data-stu-id="8ca62-135">Yes</span></span>  | <span data-ttu-id="8ca62-136">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="8ca62-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="8ca62-137">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="8ca62-137">Child elements</span></span>

|  <span data-ttu-id="8ca62-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="8ca62-138">Element</span></span> |  <span data-ttu-id="8ca62-139">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8ca62-139">Required</span></span>  |  <span data-ttu-id="8ca62-140">Описание</span><span class="sxs-lookup"><span data-stu-id="8ca62-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8ca62-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="8ca62-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="8ca62-142">Да</span><span class="sxs-lookup"><span data-stu-id="8ca62-142">Yes</span></span>   |  <span data-ttu-id="8ca62-143">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="8ca62-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="8ca62-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="8ca62-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="8ca62-145">Нет</span><span class="sxs-lookup"><span data-stu-id="8ca62-145">No</span></span>   |  <span data-ttu-id="8ca62-146">Определяет параметры для мобильного конструктивного параметра.</span><span class="sxs-lookup"><span data-stu-id="8ca62-146">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="8ca62-147">**Примечание:** Этот элемент поддерживается только в Outlook в iOS.</span><span class="sxs-lookup"><span data-stu-id="8ca62-147">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="8ca62-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="8ca62-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="8ca62-149">Нет</span><span class="sxs-lookup"><span data-stu-id="8ca62-149">No</span></span>   |  <span data-ttu-id="8ca62-150">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="8ca62-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="8ca62-151">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="8ca62-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="8ca62-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8ca62-152">xsi:type</span></span>

<span data-ttu-id="8ca62-153">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="8ca62-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="8ca62-154">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="8ca62-154">The value must be one of the following:</span></span>

- <span data-ttu-id="8ca62-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="8ca62-155">`Document` (Word)</span></span>
- <span data-ttu-id="8ca62-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="8ca62-156">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="8ca62-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="8ca62-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="8ca62-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="8ca62-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="8ca62-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="8ca62-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="8ca62-160">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="8ca62-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
