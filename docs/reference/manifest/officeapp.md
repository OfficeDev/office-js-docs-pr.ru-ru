---
title: Элемент OfficeApp в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 86f38ab77e98bb01370e40c8ada38bae171e0c2d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450459"
---
# <a name="officeapp-element"></a><span data-ttu-id="a11d2-102">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="a11d2-102">OfficeApp element</span></span>

<span data-ttu-id="a11d2-103">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="a11d2-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="a11d2-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="a11d2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a11d2-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="a11d2-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="a11d2-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="a11d2-106">Contained in</span></span>

 <span data-ttu-id="a11d2-107">_none_</span><span class="sxs-lookup"><span data-stu-id="a11d2-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="a11d2-108">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="a11d2-108">Must contain</span></span>

|<span data-ttu-id="a11d2-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="a11d2-109">**Element**</span></span>|<span data-ttu-id="a11d2-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="a11d2-110">**Content**</span></span>|<span data-ttu-id="a11d2-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="a11d2-111">**Mail**</span></span>|<span data-ttu-id="a11d2-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="a11d2-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a11d2-113">Id</span><span class="sxs-lookup"><span data-stu-id="a11d2-113">Id</span></span>](id.md)|<span data-ttu-id="a11d2-114">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-114">x</span></span>|<span data-ttu-id="a11d2-115">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-115">x</span></span>|<span data-ttu-id="a11d2-116">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-116">x</span></span>|
|[<span data-ttu-id="a11d2-117">Версия</span><span class="sxs-lookup"><span data-stu-id="a11d2-117">Version</span></span>](version.md)|<span data-ttu-id="a11d2-118">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-118">x</span></span>|<span data-ttu-id="a11d2-119">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-119">x</span></span>|<span data-ttu-id="a11d2-120">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-120">x</span></span>|
|[<span data-ttu-id="a11d2-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="a11d2-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="a11d2-122">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-122">x</span></span>|<span data-ttu-id="a11d2-123">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-123">x</span></span>|<span data-ttu-id="a11d2-124">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-124">x</span></span>|
|[<span data-ttu-id="a11d2-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="a11d2-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="a11d2-126">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-126">x</span></span>|<span data-ttu-id="a11d2-127">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-127">x</span></span>|<span data-ttu-id="a11d2-128">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-128">x</span></span>|
|[<span data-ttu-id="a11d2-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="a11d2-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="a11d2-130">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-130">x</span></span>||<span data-ttu-id="a11d2-131">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-131">x</span></span>|
|[<span data-ttu-id="a11d2-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="a11d2-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="a11d2-133">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-133">x</span></span>|<span data-ttu-id="a11d2-134">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-134">x</span></span>|<span data-ttu-id="a11d2-135">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-135">x</span></span>|
|[<span data-ttu-id="a11d2-136">Описание</span><span class="sxs-lookup"><span data-stu-id="a11d2-136">Description</span></span>](description.md)|<span data-ttu-id="a11d2-137">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-137">x</span></span>|<span data-ttu-id="a11d2-138">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-138">x</span></span>|<span data-ttu-id="a11d2-139">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-139">x</span></span>|
|[<span data-ttu-id="a11d2-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="a11d2-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="a11d2-141">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-141">x</span></span>||
|[<span data-ttu-id="a11d2-142">Разрешения</span><span class="sxs-lookup"><span data-stu-id="a11d2-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="a11d2-143">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-143">x</span></span>||<span data-ttu-id="a11d2-144">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-144">x</span></span>|
|[<span data-ttu-id="a11d2-145">Rule</span><span class="sxs-lookup"><span data-stu-id="a11d2-145">Rule</span></span>](rule.md)||<span data-ttu-id="a11d2-146">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="a11d2-147">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="a11d2-147">Can contain</span></span>

|<span data-ttu-id="a11d2-148">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="a11d2-148">**Element**</span></span>|<span data-ttu-id="a11d2-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="a11d2-149">**Content**</span></span>|<span data-ttu-id="a11d2-150">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="a11d2-150">**Mail**</span></span>|<span data-ttu-id="a11d2-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="a11d2-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a11d2-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="a11d2-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="a11d2-153">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-153">x</span></span>|<span data-ttu-id="a11d2-154">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-154">x</span></span>|<span data-ttu-id="a11d2-155">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-155">x</span></span>|
|[<span data-ttu-id="a11d2-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="a11d2-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="a11d2-157">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-157">x</span></span>|<span data-ttu-id="a11d2-158">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-158">x</span></span>|<span data-ttu-id="a11d2-159">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-159">x</span></span>|
|[<span data-ttu-id="a11d2-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="a11d2-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="a11d2-161">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-161">x</span></span>|<span data-ttu-id="a11d2-162">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-162">x</span></span>|<span data-ttu-id="a11d2-163">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-163">x</span></span>|
|[<span data-ttu-id="a11d2-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="a11d2-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="a11d2-165">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-165">x</span></span>|<span data-ttu-id="a11d2-166">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-166">x</span></span>|<span data-ttu-id="a11d2-167">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-167">x</span></span>|
|[<span data-ttu-id="a11d2-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="a11d2-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="a11d2-169">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-169">x</span></span>|<span data-ttu-id="a11d2-170">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-170">x</span></span>|<span data-ttu-id="a11d2-171">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-171">x</span></span>|
|[<span data-ttu-id="a11d2-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="a11d2-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="a11d2-173">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-173">x</span></span>|<span data-ttu-id="a11d2-174">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-174">x</span></span>|<span data-ttu-id="a11d2-175">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-175">x</span></span>|
|[<span data-ttu-id="a11d2-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="a11d2-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="a11d2-177">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-177">x</span></span>|<span data-ttu-id="a11d2-178">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-178">x</span></span>|<span data-ttu-id="a11d2-179">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-179">x</span></span>|
|[<span data-ttu-id="a11d2-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="a11d2-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="a11d2-181">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-181">x</span></span>|||
|[<span data-ttu-id="a11d2-182">Разрешения</span><span class="sxs-lookup"><span data-stu-id="a11d2-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="a11d2-183">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-183">x</span></span>||
|[<span data-ttu-id="a11d2-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="a11d2-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="a11d2-185">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-185">x</span></span>||
|[<span data-ttu-id="a11d2-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="a11d2-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="a11d2-187">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-187">x</span></span>|
|[<span data-ttu-id="a11d2-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="a11d2-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="a11d2-189">x</span><span class="sxs-lookup"><span data-stu-id="a11d2-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="a11d2-190">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a11d2-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="a11d2-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="a11d2-191">xmlns</span></span>|<span data-ttu-id="a11d2-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="a11d2-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="a11d2-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="a11d2-194">xmlns:xsi</span></span>|<span data-ttu-id="a11d2-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="a11d2-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="a11d2-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="a11d2-197">xsi:type</span></span>|<span data-ttu-id="a11d2-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="a11d2-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
