---
title: Элемент OfficeApp в файле манифеста
description: Элемент OfficeApp является корневым элементом манифеста надстройки Office.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: c5786343173d0e130df4b786f28a8689d573b6ca
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996321"
---
# <a name="officeapp-element"></a><span data-ttu-id="e657c-103">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e657c-103">OfficeApp element</span></span>

<span data-ttu-id="e657c-104">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="e657c-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="e657c-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="e657c-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e657c-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e657c-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="e657c-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e657c-107">Contained in</span></span>

 <span data-ttu-id="e657c-108">_none_</span><span class="sxs-lookup"><span data-stu-id="e657c-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="e657c-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="e657c-109">Must contain</span></span>

|<span data-ttu-id="e657c-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="e657c-110">Element</span></span>|<span data-ttu-id="e657c-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="e657c-111">Content</span></span>|<span data-ttu-id="e657c-112">Почта</span><span class="sxs-lookup"><span data-stu-id="e657c-112">Mail</span></span>|<span data-ttu-id="e657c-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="e657c-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e657c-114">Id</span><span class="sxs-lookup"><span data-stu-id="e657c-114">Id</span></span>](id.md)|<span data-ttu-id="e657c-115">x</span><span class="sxs-lookup"><span data-stu-id="e657c-115">x</span></span>|<span data-ttu-id="e657c-116">x</span><span class="sxs-lookup"><span data-stu-id="e657c-116">x</span></span>|<span data-ttu-id="e657c-117">x</span><span class="sxs-lookup"><span data-stu-id="e657c-117">x</span></span>|
|[<span data-ttu-id="e657c-118">Версия</span><span class="sxs-lookup"><span data-stu-id="e657c-118">Version</span></span>](version.md)|<span data-ttu-id="e657c-119">x</span><span class="sxs-lookup"><span data-stu-id="e657c-119">x</span></span>|<span data-ttu-id="e657c-120">x</span><span class="sxs-lookup"><span data-stu-id="e657c-120">x</span></span>|<span data-ttu-id="e657c-121">x</span><span class="sxs-lookup"><span data-stu-id="e657c-121">x</span></span>|
|[<span data-ttu-id="e657c-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="e657c-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="e657c-123">x</span><span class="sxs-lookup"><span data-stu-id="e657c-123">x</span></span>|<span data-ttu-id="e657c-124">x</span><span class="sxs-lookup"><span data-stu-id="e657c-124">x</span></span>|<span data-ttu-id="e657c-125">x</span><span class="sxs-lookup"><span data-stu-id="e657c-125">x</span></span>|
|[<span data-ttu-id="e657c-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="e657c-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="e657c-127">x</span><span class="sxs-lookup"><span data-stu-id="e657c-127">x</span></span>|<span data-ttu-id="e657c-128">x</span><span class="sxs-lookup"><span data-stu-id="e657c-128">x</span></span>|<span data-ttu-id="e657c-129">x</span><span class="sxs-lookup"><span data-stu-id="e657c-129">x</span></span>|
|[<span data-ttu-id="e657c-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="e657c-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="e657c-131">x</span><span class="sxs-lookup"><span data-stu-id="e657c-131">x</span></span>||<span data-ttu-id="e657c-132">x</span><span class="sxs-lookup"><span data-stu-id="e657c-132">x</span></span>|
|[<span data-ttu-id="e657c-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="e657c-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="e657c-134">x</span><span class="sxs-lookup"><span data-stu-id="e657c-134">x</span></span>|<span data-ttu-id="e657c-135">x</span><span class="sxs-lookup"><span data-stu-id="e657c-135">x</span></span>|<span data-ttu-id="e657c-136">x</span><span class="sxs-lookup"><span data-stu-id="e657c-136">x</span></span>|
|[<span data-ttu-id="e657c-137">Описание</span><span class="sxs-lookup"><span data-stu-id="e657c-137">Description</span></span>](description.md)|<span data-ttu-id="e657c-138">x</span><span class="sxs-lookup"><span data-stu-id="e657c-138">x</span></span>|<span data-ttu-id="e657c-139">x</span><span class="sxs-lookup"><span data-stu-id="e657c-139">x</span></span>|<span data-ttu-id="e657c-140">x</span><span class="sxs-lookup"><span data-stu-id="e657c-140">x</span></span>|
|[<span data-ttu-id="e657c-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="e657c-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="e657c-142">x</span><span class="sxs-lookup"><span data-stu-id="e657c-142">x</span></span>||
|[<span data-ttu-id="e657c-143">Разрешения</span><span class="sxs-lookup"><span data-stu-id="e657c-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="e657c-144">x</span><span class="sxs-lookup"><span data-stu-id="e657c-144">x</span></span>||<span data-ttu-id="e657c-145">x</span><span class="sxs-lookup"><span data-stu-id="e657c-145">x</span></span>|
|[<span data-ttu-id="e657c-146">Rule</span><span class="sxs-lookup"><span data-stu-id="e657c-146">Rule</span></span>](rule.md)||<span data-ttu-id="e657c-147">x</span><span class="sxs-lookup"><span data-stu-id="e657c-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="e657c-148">Может содержать</span><span class="sxs-lookup"><span data-stu-id="e657c-148">Can contain</span></span>

|<span data-ttu-id="e657c-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="e657c-149">Element</span></span>|<span data-ttu-id="e657c-150">Контентная</span><span class="sxs-lookup"><span data-stu-id="e657c-150">Content</span></span>|<span data-ttu-id="e657c-151">Почта</span><span class="sxs-lookup"><span data-stu-id="e657c-151">Mail</span></span>|<span data-ttu-id="e657c-152">Область задач</span><span class="sxs-lookup"><span data-stu-id="e657c-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="e657c-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="e657c-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="e657c-154">x</span><span class="sxs-lookup"><span data-stu-id="e657c-154">x</span></span>|<span data-ttu-id="e657c-155">x</span><span class="sxs-lookup"><span data-stu-id="e657c-155">x</span></span>|<span data-ttu-id="e657c-156">x</span><span class="sxs-lookup"><span data-stu-id="e657c-156">x</span></span>|
|[<span data-ttu-id="e657c-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="e657c-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="e657c-158">x</span><span class="sxs-lookup"><span data-stu-id="e657c-158">x</span></span>|<span data-ttu-id="e657c-159">x</span><span class="sxs-lookup"><span data-stu-id="e657c-159">x</span></span>|<span data-ttu-id="e657c-160">x</span><span class="sxs-lookup"><span data-stu-id="e657c-160">x</span></span>|
|[<span data-ttu-id="e657c-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="e657c-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="e657c-162">x</span><span class="sxs-lookup"><span data-stu-id="e657c-162">x</span></span>|<span data-ttu-id="e657c-163">x</span><span class="sxs-lookup"><span data-stu-id="e657c-163">x</span></span>|<span data-ttu-id="e657c-164">x</span><span class="sxs-lookup"><span data-stu-id="e657c-164">x</span></span>|
|[<span data-ttu-id="e657c-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="e657c-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="e657c-166">x</span><span class="sxs-lookup"><span data-stu-id="e657c-166">x</span></span>|<span data-ttu-id="e657c-167">x</span><span class="sxs-lookup"><span data-stu-id="e657c-167">x</span></span>|<span data-ttu-id="e657c-168">x</span><span class="sxs-lookup"><span data-stu-id="e657c-168">x</span></span>|
|[<span data-ttu-id="e657c-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="e657c-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="e657c-170">x</span><span class="sxs-lookup"><span data-stu-id="e657c-170">x</span></span>|<span data-ttu-id="e657c-171">x</span><span class="sxs-lookup"><span data-stu-id="e657c-171">x</span></span>|<span data-ttu-id="e657c-172">x</span><span class="sxs-lookup"><span data-stu-id="e657c-172">x</span></span>|
|[<span data-ttu-id="e657c-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="e657c-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="e657c-174">x</span><span class="sxs-lookup"><span data-stu-id="e657c-174">x</span></span>|<span data-ttu-id="e657c-175">x</span><span class="sxs-lookup"><span data-stu-id="e657c-175">x</span></span>|<span data-ttu-id="e657c-176">x</span><span class="sxs-lookup"><span data-stu-id="e657c-176">x</span></span>|
|[<span data-ttu-id="e657c-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="e657c-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="e657c-178">x</span><span class="sxs-lookup"><span data-stu-id="e657c-178">x</span></span>|<span data-ttu-id="e657c-179">x</span><span class="sxs-lookup"><span data-stu-id="e657c-179">x</span></span>|<span data-ttu-id="e657c-180">x</span><span class="sxs-lookup"><span data-stu-id="e657c-180">x</span></span>|
|[<span data-ttu-id="e657c-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="e657c-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="e657c-182">x</span><span class="sxs-lookup"><span data-stu-id="e657c-182">x</span></span>|||
|[<span data-ttu-id="e657c-183">Разрешения</span><span class="sxs-lookup"><span data-stu-id="e657c-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="e657c-184">x</span><span class="sxs-lookup"><span data-stu-id="e657c-184">x</span></span>||
|[<span data-ttu-id="e657c-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="e657c-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="e657c-186">x</span><span class="sxs-lookup"><span data-stu-id="e657c-186">x</span></span>||
|[<span data-ttu-id="e657c-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="e657c-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="e657c-188">x</span><span class="sxs-lookup"><span data-stu-id="e657c-188">x</span></span>|
|[<span data-ttu-id="e657c-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="e657c-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="e657c-190">x</span><span class="sxs-lookup"><span data-stu-id="e657c-190">x</span></span>|<span data-ttu-id="e657c-191">x</span><span class="sxs-lookup"><span data-stu-id="e657c-191">x</span></span>|<span data-ttu-id="e657c-192">x</span><span class="sxs-lookup"><span data-stu-id="e657c-192">x</span></span>|
|[<span data-ttu-id="e657c-193">екстендедоверридес</span><span class="sxs-lookup"><span data-stu-id="e657c-193">ExtendedOverrides</span></span>](extendedoverrides.md)|||<span data-ttu-id="e657c-194">x</span><span class="sxs-lookup"><span data-stu-id="e657c-194">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="e657c-195">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e657c-195">Attributes</span></span>

|<span data-ttu-id="e657c-196">Атрибут</span><span class="sxs-lookup"><span data-stu-id="e657c-196">Attribute</span></span>|<span data-ttu-id="e657c-197">Описание</span><span class="sxs-lookup"><span data-stu-id="e657c-197">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="e657c-198">xmlns</span><span class="sxs-lookup"><span data-stu-id="e657c-198">xmlns</span></span>|<span data-ttu-id="e657c-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="e657c-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="e657c-201">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="e657c-201">xmlns:xsi</span></span>|<span data-ttu-id="e657c-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="e657c-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="e657c-204">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e657c-204">xsi:type</span></span>|<span data-ttu-id="e657c-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="e657c-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
