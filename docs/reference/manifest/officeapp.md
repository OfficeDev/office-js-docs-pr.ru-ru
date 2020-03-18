---
title: Элемент OfficeApp в файле манифеста
description: Элемент OfficeApp является корневым элементом манифеста надстройки Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 038933f2d06ee5f485dbdb7dd7abdbd95fb97c7d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720598"
---
# <a name="officeapp-element"></a><span data-ttu-id="dd566-103">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="dd566-103">OfficeApp element</span></span>

<span data-ttu-id="dd566-104">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="dd566-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="dd566-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="dd566-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="dd566-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="dd566-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="dd566-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="dd566-107">Contained in</span></span>

 <span data-ttu-id="dd566-108">_none_</span><span class="sxs-lookup"><span data-stu-id="dd566-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="dd566-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="dd566-109">Must contain</span></span>

|<span data-ttu-id="dd566-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="dd566-110">**Element**</span></span>|<span data-ttu-id="dd566-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="dd566-111">**Content**</span></span>|<span data-ttu-id="dd566-112">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="dd566-112">**Mail**</span></span>|<span data-ttu-id="dd566-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="dd566-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="dd566-114">Id</span><span class="sxs-lookup"><span data-stu-id="dd566-114">Id</span></span>](id.md)|<span data-ttu-id="dd566-115">x</span><span class="sxs-lookup"><span data-stu-id="dd566-115">x</span></span>|<span data-ttu-id="dd566-116">x</span><span class="sxs-lookup"><span data-stu-id="dd566-116">x</span></span>|<span data-ttu-id="dd566-117">x</span><span class="sxs-lookup"><span data-stu-id="dd566-117">x</span></span>|
|[<span data-ttu-id="dd566-118">Версия</span><span class="sxs-lookup"><span data-stu-id="dd566-118">Version</span></span>](version.md)|<span data-ttu-id="dd566-119">x</span><span class="sxs-lookup"><span data-stu-id="dd566-119">x</span></span>|<span data-ttu-id="dd566-120">x</span><span class="sxs-lookup"><span data-stu-id="dd566-120">x</span></span>|<span data-ttu-id="dd566-121">x</span><span class="sxs-lookup"><span data-stu-id="dd566-121">x</span></span>|
|[<span data-ttu-id="dd566-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="dd566-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="dd566-123">x</span><span class="sxs-lookup"><span data-stu-id="dd566-123">x</span></span>|<span data-ttu-id="dd566-124">x</span><span class="sxs-lookup"><span data-stu-id="dd566-124">x</span></span>|<span data-ttu-id="dd566-125">x</span><span class="sxs-lookup"><span data-stu-id="dd566-125">x</span></span>|
|[<span data-ttu-id="dd566-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="dd566-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="dd566-127">x</span><span class="sxs-lookup"><span data-stu-id="dd566-127">x</span></span>|<span data-ttu-id="dd566-128">x</span><span class="sxs-lookup"><span data-stu-id="dd566-128">x</span></span>|<span data-ttu-id="dd566-129">x</span><span class="sxs-lookup"><span data-stu-id="dd566-129">x</span></span>|
|[<span data-ttu-id="dd566-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="dd566-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="dd566-131">x</span><span class="sxs-lookup"><span data-stu-id="dd566-131">x</span></span>||<span data-ttu-id="dd566-132">x</span><span class="sxs-lookup"><span data-stu-id="dd566-132">x</span></span>|
|[<span data-ttu-id="dd566-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="dd566-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="dd566-134">x</span><span class="sxs-lookup"><span data-stu-id="dd566-134">x</span></span>|<span data-ttu-id="dd566-135">x</span><span class="sxs-lookup"><span data-stu-id="dd566-135">x</span></span>|<span data-ttu-id="dd566-136">x</span><span class="sxs-lookup"><span data-stu-id="dd566-136">x</span></span>|
|[<span data-ttu-id="dd566-137">Описание</span><span class="sxs-lookup"><span data-stu-id="dd566-137">Description</span></span>](description.md)|<span data-ttu-id="dd566-138">x</span><span class="sxs-lookup"><span data-stu-id="dd566-138">x</span></span>|<span data-ttu-id="dd566-139">x</span><span class="sxs-lookup"><span data-stu-id="dd566-139">x</span></span>|<span data-ttu-id="dd566-140">x</span><span class="sxs-lookup"><span data-stu-id="dd566-140">x</span></span>|
|[<span data-ttu-id="dd566-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="dd566-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="dd566-142">x</span><span class="sxs-lookup"><span data-stu-id="dd566-142">x</span></span>||
|[<span data-ttu-id="dd566-143">Разрешения</span><span class="sxs-lookup"><span data-stu-id="dd566-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="dd566-144">x</span><span class="sxs-lookup"><span data-stu-id="dd566-144">x</span></span>||<span data-ttu-id="dd566-145">x</span><span class="sxs-lookup"><span data-stu-id="dd566-145">x</span></span>|
|[<span data-ttu-id="dd566-146">Rule</span><span class="sxs-lookup"><span data-stu-id="dd566-146">Rule</span></span>](rule.md)||<span data-ttu-id="dd566-147">x</span><span class="sxs-lookup"><span data-stu-id="dd566-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="dd566-148">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="dd566-148">Can contain</span></span>

|<span data-ttu-id="dd566-149">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="dd566-149">**Element**</span></span>|<span data-ttu-id="dd566-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="dd566-150">**Content**</span></span>|<span data-ttu-id="dd566-151">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="dd566-151">**Mail**</span></span>|<span data-ttu-id="dd566-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="dd566-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="dd566-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="dd566-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="dd566-154">x</span><span class="sxs-lookup"><span data-stu-id="dd566-154">x</span></span>|<span data-ttu-id="dd566-155">x</span><span class="sxs-lookup"><span data-stu-id="dd566-155">x</span></span>|<span data-ttu-id="dd566-156">x</span><span class="sxs-lookup"><span data-stu-id="dd566-156">x</span></span>|
|[<span data-ttu-id="dd566-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="dd566-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="dd566-158">x</span><span class="sxs-lookup"><span data-stu-id="dd566-158">x</span></span>|<span data-ttu-id="dd566-159">x</span><span class="sxs-lookup"><span data-stu-id="dd566-159">x</span></span>|<span data-ttu-id="dd566-160">x</span><span class="sxs-lookup"><span data-stu-id="dd566-160">x</span></span>|
|[<span data-ttu-id="dd566-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="dd566-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="dd566-162">x</span><span class="sxs-lookup"><span data-stu-id="dd566-162">x</span></span>|<span data-ttu-id="dd566-163">x</span><span class="sxs-lookup"><span data-stu-id="dd566-163">x</span></span>|<span data-ttu-id="dd566-164">x</span><span class="sxs-lookup"><span data-stu-id="dd566-164">x</span></span>|
|[<span data-ttu-id="dd566-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="dd566-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="dd566-166">x</span><span class="sxs-lookup"><span data-stu-id="dd566-166">x</span></span>|<span data-ttu-id="dd566-167">x</span><span class="sxs-lookup"><span data-stu-id="dd566-167">x</span></span>|<span data-ttu-id="dd566-168">x</span><span class="sxs-lookup"><span data-stu-id="dd566-168">x</span></span>|
|[<span data-ttu-id="dd566-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="dd566-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="dd566-170">x</span><span class="sxs-lookup"><span data-stu-id="dd566-170">x</span></span>|<span data-ttu-id="dd566-171">x</span><span class="sxs-lookup"><span data-stu-id="dd566-171">x</span></span>|<span data-ttu-id="dd566-172">x</span><span class="sxs-lookup"><span data-stu-id="dd566-172">x</span></span>|
|[<span data-ttu-id="dd566-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="dd566-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="dd566-174">x</span><span class="sxs-lookup"><span data-stu-id="dd566-174">x</span></span>|<span data-ttu-id="dd566-175">x</span><span class="sxs-lookup"><span data-stu-id="dd566-175">x</span></span>|<span data-ttu-id="dd566-176">x</span><span class="sxs-lookup"><span data-stu-id="dd566-176">x</span></span>|
|[<span data-ttu-id="dd566-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd566-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="dd566-178">x</span><span class="sxs-lookup"><span data-stu-id="dd566-178">x</span></span>|<span data-ttu-id="dd566-179">x</span><span class="sxs-lookup"><span data-stu-id="dd566-179">x</span></span>|<span data-ttu-id="dd566-180">x</span><span class="sxs-lookup"><span data-stu-id="dd566-180">x</span></span>|
|[<span data-ttu-id="dd566-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="dd566-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="dd566-182">x</span><span class="sxs-lookup"><span data-stu-id="dd566-182">x</span></span>|||
|[<span data-ttu-id="dd566-183">Разрешения</span><span class="sxs-lookup"><span data-stu-id="dd566-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="dd566-184">x</span><span class="sxs-lookup"><span data-stu-id="dd566-184">x</span></span>||
|[<span data-ttu-id="dd566-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="dd566-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="dd566-186">x</span><span class="sxs-lookup"><span data-stu-id="dd566-186">x</span></span>||
|[<span data-ttu-id="dd566-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="dd566-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="dd566-188">x</span><span class="sxs-lookup"><span data-stu-id="dd566-188">x</span></span>|
|[<span data-ttu-id="dd566-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="dd566-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="dd566-190">x</span><span class="sxs-lookup"><span data-stu-id="dd566-190">x</span></span>|<span data-ttu-id="dd566-191">x</span><span class="sxs-lookup"><span data-stu-id="dd566-191">x</span></span>|<span data-ttu-id="dd566-192">x</span><span class="sxs-lookup"><span data-stu-id="dd566-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="dd566-193">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd566-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="dd566-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="dd566-194">xmlns</span></span>|<span data-ttu-id="dd566-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="dd566-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="dd566-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="dd566-197">xmlns:xsi</span></span>|<span data-ttu-id="dd566-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="dd566-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="dd566-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="dd566-200">xsi:type</span></span>|<span data-ttu-id="dd566-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="dd566-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
