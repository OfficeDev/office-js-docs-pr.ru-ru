---
title: Элемент OfficeApp в файле манифеста
description: Элемент OfficeApp является корневым элементом манифеста надстройки Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: b6f3102a97794a19366b06734789e01fc4bc4f9d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611528"
---
# <a name="officeapp-element"></a><span data-ttu-id="0c31e-103">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0c31e-103">OfficeApp element</span></span>

<span data-ttu-id="0c31e-104">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="0c31e-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="0c31e-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="0c31e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0c31e-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0c31e-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="0c31e-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="0c31e-107">Contained in</span></span>

 <span data-ttu-id="0c31e-108">_none_</span><span class="sxs-lookup"><span data-stu-id="0c31e-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="0c31e-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="0c31e-109">Must contain</span></span>

|<span data-ttu-id="0c31e-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="0c31e-110">**Element**</span></span>|<span data-ttu-id="0c31e-111">**Content**</span><span class="sxs-lookup"><span data-stu-id="0c31e-111">**Content**</span></span>|<span data-ttu-id="0c31e-112">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="0c31e-112">**Mail**</span></span>|<span data-ttu-id="0c31e-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="0c31e-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0c31e-114">Id</span><span class="sxs-lookup"><span data-stu-id="0c31e-114">Id</span></span>](id.md)|<span data-ttu-id="0c31e-115">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-115">x</span></span>|<span data-ttu-id="0c31e-116">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-116">x</span></span>|<span data-ttu-id="0c31e-117">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-117">x</span></span>|
|[<span data-ttu-id="0c31e-118">Версия</span><span class="sxs-lookup"><span data-stu-id="0c31e-118">Version</span></span>](version.md)|<span data-ttu-id="0c31e-119">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-119">x</span></span>|<span data-ttu-id="0c31e-120">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-120">x</span></span>|<span data-ttu-id="0c31e-121">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-121">x</span></span>|
|[<span data-ttu-id="0c31e-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="0c31e-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="0c31e-123">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-123">x</span></span>|<span data-ttu-id="0c31e-124">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-124">x</span></span>|<span data-ttu-id="0c31e-125">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-125">x</span></span>|
|[<span data-ttu-id="0c31e-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="0c31e-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="0c31e-127">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-127">x</span></span>|<span data-ttu-id="0c31e-128">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-128">x</span></span>|<span data-ttu-id="0c31e-129">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-129">x</span></span>|
|[<span data-ttu-id="0c31e-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="0c31e-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="0c31e-131">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-131">x</span></span>||<span data-ttu-id="0c31e-132">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-132">x</span></span>|
|[<span data-ttu-id="0c31e-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="0c31e-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="0c31e-134">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-134">x</span></span>|<span data-ttu-id="0c31e-135">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-135">x</span></span>|<span data-ttu-id="0c31e-136">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-136">x</span></span>|
|[<span data-ttu-id="0c31e-137">Описание</span><span class="sxs-lookup"><span data-stu-id="0c31e-137">Description</span></span>](description.md)|<span data-ttu-id="0c31e-138">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-138">x</span></span>|<span data-ttu-id="0c31e-139">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-139">x</span></span>|<span data-ttu-id="0c31e-140">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-140">x</span></span>|
|[<span data-ttu-id="0c31e-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="0c31e-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="0c31e-142">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-142">x</span></span>||
|[<span data-ttu-id="0c31e-143">Разрешения</span><span class="sxs-lookup"><span data-stu-id="0c31e-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="0c31e-144">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-144">x</span></span>||<span data-ttu-id="0c31e-145">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-145">x</span></span>|
|[<span data-ttu-id="0c31e-146">Rule</span><span class="sxs-lookup"><span data-stu-id="0c31e-146">Rule</span></span>](rule.md)||<span data-ttu-id="0c31e-147">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="0c31e-148">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="0c31e-148">Can contain</span></span>

|<span data-ttu-id="0c31e-149">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="0c31e-149">**Element**</span></span>|<span data-ttu-id="0c31e-150">**Content**</span><span class="sxs-lookup"><span data-stu-id="0c31e-150">**Content**</span></span>|<span data-ttu-id="0c31e-151">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="0c31e-151">**Mail**</span></span>|<span data-ttu-id="0c31e-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="0c31e-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0c31e-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="0c31e-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="0c31e-154">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-154">x</span></span>|<span data-ttu-id="0c31e-155">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-155">x</span></span>|<span data-ttu-id="0c31e-156">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-156">x</span></span>|
|[<span data-ttu-id="0c31e-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="0c31e-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="0c31e-158">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-158">x</span></span>|<span data-ttu-id="0c31e-159">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-159">x</span></span>|<span data-ttu-id="0c31e-160">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-160">x</span></span>|
|[<span data-ttu-id="0c31e-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="0c31e-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="0c31e-162">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-162">x</span></span>|<span data-ttu-id="0c31e-163">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-163">x</span></span>|<span data-ttu-id="0c31e-164">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-164">x</span></span>|
|[<span data-ttu-id="0c31e-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="0c31e-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="0c31e-166">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-166">x</span></span>|<span data-ttu-id="0c31e-167">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-167">x</span></span>|<span data-ttu-id="0c31e-168">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-168">x</span></span>|
|[<span data-ttu-id="0c31e-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="0c31e-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="0c31e-170">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-170">x</span></span>|<span data-ttu-id="0c31e-171">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-171">x</span></span>|<span data-ttu-id="0c31e-172">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-172">x</span></span>|
|[<span data-ttu-id="0c31e-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="0c31e-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="0c31e-174">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-174">x</span></span>|<span data-ttu-id="0c31e-175">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-175">x</span></span>|<span data-ttu-id="0c31e-176">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-176">x</span></span>|
|[<span data-ttu-id="0c31e-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="0c31e-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="0c31e-178">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-178">x</span></span>|<span data-ttu-id="0c31e-179">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-179">x</span></span>|<span data-ttu-id="0c31e-180">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-180">x</span></span>|
|[<span data-ttu-id="0c31e-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="0c31e-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="0c31e-182">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-182">x</span></span>|||
|[<span data-ttu-id="0c31e-183">Разрешения</span><span class="sxs-lookup"><span data-stu-id="0c31e-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="0c31e-184">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-184">x</span></span>||
|[<span data-ttu-id="0c31e-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="0c31e-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="0c31e-186">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-186">x</span></span>||
|[<span data-ttu-id="0c31e-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="0c31e-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="0c31e-188">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-188">x</span></span>|
|[<span data-ttu-id="0c31e-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="0c31e-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="0c31e-190">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-190">x</span></span>|<span data-ttu-id="0c31e-191">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-191">x</span></span>|<span data-ttu-id="0c31e-192">x</span><span class="sxs-lookup"><span data-stu-id="0c31e-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="0c31e-193">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0c31e-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="0c31e-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="0c31e-194">xmlns</span></span>|<span data-ttu-id="0c31e-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="0c31e-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="0c31e-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="0c31e-197">xmlns:xsi</span></span>|<span data-ttu-id="0c31e-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="0c31e-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="0c31e-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0c31e-200">xsi:type</span></span>|<span data-ttu-id="0c31e-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="0c31e-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
