---
title: Элемент Host в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 37b772261ad82b4f899e73314a08ffd1dd03b442
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432811"
---
# <a name="host-element"></a><span data-ttu-id="9b92d-102">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="9b92d-102">Host element</span></span>

<span data-ttu-id="9b92d-103">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="9b92d-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="9b92d-104">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="9b92d-104">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="9b92d-105">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="9b92d-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="9b92d-106">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="9b92d-106">Basic manifest</span></span>

<span data-ttu-id="9b92d-107">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="9b92d-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="9b92d-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9b92d-108">Attributes</span></span>

| <span data-ttu-id="9b92d-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="9b92d-109">Attribute</span></span>     | <span data-ttu-id="9b92d-110">Тип</span><span class="sxs-lookup"><span data-stu-id="9b92d-110">Type</span></span>   | <span data-ttu-id="9b92d-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9b92d-111">Required</span></span> | <span data-ttu-id="9b92d-112">Описание</span><span class="sxs-lookup"><span data-stu-id="9b92d-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="9b92d-113">Name</span><span class="sxs-lookup"><span data-stu-id="9b92d-113">Name</span></span>](#name) | <span data-ttu-id="9b92d-114">string</span><span class="sxs-lookup"><span data-stu-id="9b92d-114">string</span></span> | <span data-ttu-id="9b92d-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9b92d-115">required</span></span> | <span data-ttu-id="9b92d-116">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="9b92d-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="9b92d-117">Имя</span><span class="sxs-lookup"><span data-stu-id="9b92d-117">Name</span></span>
<span data-ttu-id="9b92d-p102">Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="9b92d-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="9b92d-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9b92d-120">`Document` (Word)</span></span>
- <span data-ttu-id="9b92d-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="9b92d-121">`Database` (Access)</span></span>
- <span data-ttu-id="9b92d-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9b92d-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="9b92d-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9b92d-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9b92d-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9b92d-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9b92d-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="9b92d-125">`Project` (Project)</span></span>
- <span data-ttu-id="9b92d-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9b92d-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="9b92d-127">Пример</span><span class="sxs-lookup"><span data-stu-id="9b92d-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="9b92d-128">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="9b92d-128">VersionOverrides node</span></span>
<span data-ttu-id="9b92d-129">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="9b92d-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="9b92d-130">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9b92d-130">Attributes</span></span>

|  <span data-ttu-id="9b92d-131">Атрибут</span><span class="sxs-lookup"><span data-stu-id="9b92d-131">Attribute</span></span>  |  <span data-ttu-id="9b92d-132">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9b92d-132">Required</span></span>  |  <span data-ttu-id="9b92d-133">Описание</span><span class="sxs-lookup"><span data-stu-id="9b92d-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9b92d-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9b92d-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="9b92d-135">Да</span><span class="sxs-lookup"><span data-stu-id="9b92d-135">Yes</span></span>  | <span data-ttu-id="9b92d-136">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="9b92d-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="9b92d-137">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9b92d-137">Child elements</span></span>

|  <span data-ttu-id="9b92d-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="9b92d-138">Element</span></span> |  <span data-ttu-id="9b92d-139">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9b92d-139">Required</span></span>  |  <span data-ttu-id="9b92d-140">Описание</span><span class="sxs-lookup"><span data-stu-id="9b92d-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9b92d-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="9b92d-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="9b92d-142">Да</span><span class="sxs-lookup"><span data-stu-id="9b92d-142">Yes</span></span>   |  <span data-ttu-id="9b92d-143">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="9b92d-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="9b92d-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="9b92d-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="9b92d-145">Нет</span><span class="sxs-lookup"><span data-stu-id="9b92d-145">No</span></span>   |  <span data-ttu-id="9b92d-p103">Определяет параметры форм-фактора мобильного устройства. **Примечание.** Этот элемент поддерживается только в Outlook для iOS.</span><span class="sxs-lookup"><span data-stu-id="9b92d-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="9b92d-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="9b92d-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="9b92d-149">Нет</span><span class="sxs-lookup"><span data-stu-id="9b92d-149">No</span></span>   |  <span data-ttu-id="9b92d-150">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="9b92d-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="9b92d-151">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="9b92d-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="9b92d-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="9b92d-152">xsi:type</span></span>

<span data-ttu-id="9b92d-153">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="9b92d-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="9b92d-154">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="9b92d-154">The value must be one of the following:</span></span>

- <span data-ttu-id="9b92d-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="9b92d-155">`Document` (Word)</span></span>
- <span data-ttu-id="9b92d-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="9b92d-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="9b92d-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="9b92d-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="9b92d-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="9b92d-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="9b92d-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="9b92d-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="9b92d-160">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="9b92d-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
