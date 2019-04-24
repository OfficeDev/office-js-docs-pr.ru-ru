---
title: Элемент Host в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f496e3e0c16f24d20e1d1db76208e61267235131
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450508"
---
# <a name="host-element"></a><span data-ttu-id="1bb99-102">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="1bb99-102">Host element</span></span>

<span data-ttu-id="1bb99-103">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="1bb99-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="1bb99-104">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="1bb99-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="1bb99-105">Функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="1bb99-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="1bb99-106">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="1bb99-106">Basic manifest</span></span>

<span data-ttu-id="1bb99-107">Если ведущее приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяет атрибут `Name`.</span><span class="sxs-lookup"><span data-stu-id="1bb99-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="1bb99-108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1bb99-108">Attributes</span></span>

| <span data-ttu-id="1bb99-109">Атрибут</span><span class="sxs-lookup"><span data-stu-id="1bb99-109">Attribute</span></span>     | <span data-ttu-id="1bb99-110">Тип</span><span class="sxs-lookup"><span data-stu-id="1bb99-110">Type</span></span>   | <span data-ttu-id="1bb99-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1bb99-111">Required</span></span> | <span data-ttu-id="1bb99-112">Описание</span><span class="sxs-lookup"><span data-stu-id="1bb99-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="1bb99-113">Name</span><span class="sxs-lookup"><span data-stu-id="1bb99-113">Name</span></span>](#name) | <span data-ttu-id="1bb99-114">string</span><span class="sxs-lookup"><span data-stu-id="1bb99-114">string</span></span> | <span data-ttu-id="1bb99-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1bb99-115">required</span></span> | <span data-ttu-id="1bb99-116">Имя типа ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="1bb99-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="1bb99-117">Имя</span><span class="sxs-lookup"><span data-stu-id="1bb99-117">Name</span></span>
<span data-ttu-id="1bb99-p102">Определяет тип ведущего приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="1bb99-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="1bb99-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="1bb99-120">`Document` (Word)</span></span>
- <span data-ttu-id="1bb99-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="1bb99-121">`Database` (Access)</span></span>
- <span data-ttu-id="1bb99-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="1bb99-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="1bb99-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="1bb99-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="1bb99-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="1bb99-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="1bb99-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="1bb99-125">`Project` (Project)</span></span>
- <span data-ttu-id="1bb99-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="1bb99-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="1bb99-127">Пример</span><span class="sxs-lookup"><span data-stu-id="1bb99-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="1bb99-128">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="1bb99-128">VersionOverrides node</span></span>
<span data-ttu-id="1bb99-129">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="1bb99-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="1bb99-130">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1bb99-130">Attributes</span></span>

|  <span data-ttu-id="1bb99-131">Атрибут</span><span class="sxs-lookup"><span data-stu-id="1bb99-131">Attribute</span></span>  |  <span data-ttu-id="1bb99-132">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1bb99-132">Required</span></span>  |  <span data-ttu-id="1bb99-133">Описание</span><span class="sxs-lookup"><span data-stu-id="1bb99-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1bb99-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1bb99-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="1bb99-135">Да</span><span class="sxs-lookup"><span data-stu-id="1bb99-135">Yes</span></span>  | <span data-ttu-id="1bb99-136">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="1bb99-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="1bb99-137">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="1bb99-137">Child elements</span></span>

|  <span data-ttu-id="1bb99-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="1bb99-138">Element</span></span> |  <span data-ttu-id="1bb99-139">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1bb99-139">Required</span></span>  |  <span data-ttu-id="1bb99-140">Описание</span><span class="sxs-lookup"><span data-stu-id="1bb99-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1bb99-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="1bb99-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="1bb99-142">Да</span><span class="sxs-lookup"><span data-stu-id="1bb99-142">Yes</span></span>   |  <span data-ttu-id="1bb99-143">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="1bb99-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="1bb99-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="1bb99-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="1bb99-145">Нет</span><span class="sxs-lookup"><span data-stu-id="1bb99-145">No</span></span>   |  <span data-ttu-id="1bb99-p103">Определяет параметры форм-фактора мобильного устройства. **Примечание.** Этот элемент поддерживается только в Outlook для iOS.</span><span class="sxs-lookup"><span data-stu-id="1bb99-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="1bb99-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="1bb99-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="1bb99-149">Нет</span><span class="sxs-lookup"><span data-stu-id="1bb99-149">No</span></span>   |  <span data-ttu-id="1bb99-150">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="1bb99-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="1bb99-151">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="1bb99-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="1bb99-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1bb99-152">xsi:type</span></span>

<span data-ttu-id="1bb99-153">Указывает, к какому ведущему приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="1bb99-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="1bb99-154">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="1bb99-154">The value must be one of the following:</span></span>

- <span data-ttu-id="1bb99-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="1bb99-155">`Document` (Word)</span></span>
- <span data-ttu-id="1bb99-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="1bb99-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="1bb99-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="1bb99-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="1bb99-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="1bb99-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="1bb99-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="1bb99-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="1bb99-160">Пример ведущего приложения</span><span class="sxs-lookup"><span data-stu-id="1bb99-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
