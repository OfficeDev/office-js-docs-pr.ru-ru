---
title: Элемент OfficeApp в файле манифеста
description: Элемент OfficeApp является корневым элементом манифеста надстройки Office.
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 770c764db6d8d7d1d2e870e48437de7c8f887101
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641461"
---
# <a name="officeapp-element"></a><span data-ttu-id="bb4fa-103">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="bb4fa-103">OfficeApp element</span></span>

<span data-ttu-id="bb4fa-104">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="bb4fa-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="bb4fa-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="bb4fa-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bb4fa-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="bb4fa-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="bb4fa-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="bb4fa-107">Contained in</span></span>

 <span data-ttu-id="bb4fa-108">_none_</span><span class="sxs-lookup"><span data-stu-id="bb4fa-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="bb4fa-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="bb4fa-109">Must contain</span></span>

|<span data-ttu-id="bb4fa-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="bb4fa-110">Element</span></span>|<span data-ttu-id="bb4fa-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="bb4fa-111">Content</span></span>|<span data-ttu-id="bb4fa-112">Почта</span><span class="sxs-lookup"><span data-stu-id="bb4fa-112">Mail</span></span>|<span data-ttu-id="bb4fa-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="bb4fa-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="bb4fa-114">Id</span><span class="sxs-lookup"><span data-stu-id="bb4fa-114">Id</span></span>](id.md)|<span data-ttu-id="bb4fa-115">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-115">x</span></span>|<span data-ttu-id="bb4fa-116">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-116">x</span></span>|<span data-ttu-id="bb4fa-117">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-117">x</span></span>|
|[<span data-ttu-id="bb4fa-118">Версия</span><span class="sxs-lookup"><span data-stu-id="bb4fa-118">Version</span></span>](version.md)|<span data-ttu-id="bb4fa-119">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-119">x</span></span>|<span data-ttu-id="bb4fa-120">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-120">x</span></span>|<span data-ttu-id="bb4fa-121">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-121">x</span></span>|
|[<span data-ttu-id="bb4fa-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="bb4fa-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="bb4fa-123">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-123">x</span></span>|<span data-ttu-id="bb4fa-124">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-124">x</span></span>|<span data-ttu-id="bb4fa-125">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-125">x</span></span>|
|[<span data-ttu-id="bb4fa-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="bb4fa-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="bb4fa-127">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-127">x</span></span>|<span data-ttu-id="bb4fa-128">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-128">x</span></span>|<span data-ttu-id="bb4fa-129">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-129">x</span></span>|
|[<span data-ttu-id="bb4fa-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="bb4fa-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="bb4fa-131">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-131">x</span></span>||<span data-ttu-id="bb4fa-132">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-132">x</span></span>|
|[<span data-ttu-id="bb4fa-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="bb4fa-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="bb4fa-134">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-134">x</span></span>|<span data-ttu-id="bb4fa-135">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-135">x</span></span>|<span data-ttu-id="bb4fa-136">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-136">x</span></span>|
|[<span data-ttu-id="bb4fa-137">Описание</span><span class="sxs-lookup"><span data-stu-id="bb4fa-137">Description</span></span>](description.md)|<span data-ttu-id="bb4fa-138">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-138">x</span></span>|<span data-ttu-id="bb4fa-139">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-139">x</span></span>|<span data-ttu-id="bb4fa-140">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-140">x</span></span>|
|[<span data-ttu-id="bb4fa-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="bb4fa-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="bb4fa-142">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-142">x</span></span>||
|[<span data-ttu-id="bb4fa-143">Разрешения</span><span class="sxs-lookup"><span data-stu-id="bb4fa-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="bb4fa-144">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-144">x</span></span>||<span data-ttu-id="bb4fa-145">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-145">x</span></span>|
|[<span data-ttu-id="bb4fa-146">Rule</span><span class="sxs-lookup"><span data-stu-id="bb4fa-146">Rule</span></span>](rule.md)||<span data-ttu-id="bb4fa-147">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="bb4fa-148">Может содержать</span><span class="sxs-lookup"><span data-stu-id="bb4fa-148">Can contain</span></span>

|<span data-ttu-id="bb4fa-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="bb4fa-149">Element</span></span>|<span data-ttu-id="bb4fa-150">Контентная</span><span class="sxs-lookup"><span data-stu-id="bb4fa-150">Content</span></span>|<span data-ttu-id="bb4fa-151">Почта</span><span class="sxs-lookup"><span data-stu-id="bb4fa-151">Mail</span></span>|<span data-ttu-id="bb4fa-152">Область задач</span><span class="sxs-lookup"><span data-stu-id="bb4fa-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="bb4fa-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="bb4fa-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="bb4fa-154">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-154">x</span></span>|<span data-ttu-id="bb4fa-155">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-155">x</span></span>|<span data-ttu-id="bb4fa-156">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-156">x</span></span>|
|[<span data-ttu-id="bb4fa-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="bb4fa-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="bb4fa-158">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-158">x</span></span>|<span data-ttu-id="bb4fa-159">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-159">x</span></span>|<span data-ttu-id="bb4fa-160">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-160">x</span></span>|
|[<span data-ttu-id="bb4fa-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="bb4fa-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="bb4fa-162">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-162">x</span></span>|<span data-ttu-id="bb4fa-163">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-163">x</span></span>|<span data-ttu-id="bb4fa-164">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-164">x</span></span>|
|[<span data-ttu-id="bb4fa-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="bb4fa-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="bb4fa-166">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-166">x</span></span>|<span data-ttu-id="bb4fa-167">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-167">x</span></span>|<span data-ttu-id="bb4fa-168">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-168">x</span></span>|
|[<span data-ttu-id="bb4fa-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="bb4fa-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="bb4fa-170">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-170">x</span></span>|<span data-ttu-id="bb4fa-171">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-171">x</span></span>|<span data-ttu-id="bb4fa-172">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-172">x</span></span>|
|[<span data-ttu-id="bb4fa-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="bb4fa-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="bb4fa-174">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-174">x</span></span>|<span data-ttu-id="bb4fa-175">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-175">x</span></span>|<span data-ttu-id="bb4fa-176">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-176">x</span></span>|
|[<span data-ttu-id="bb4fa-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="bb4fa-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="bb4fa-178">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-178">x</span></span>|<span data-ttu-id="bb4fa-179">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-179">x</span></span>|<span data-ttu-id="bb4fa-180">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-180">x</span></span>|
|[<span data-ttu-id="bb4fa-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="bb4fa-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="bb4fa-182">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-182">x</span></span>|||
|[<span data-ttu-id="bb4fa-183">Разрешения</span><span class="sxs-lookup"><span data-stu-id="bb4fa-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="bb4fa-184">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-184">x</span></span>||
|[<span data-ttu-id="bb4fa-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="bb4fa-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="bb4fa-186">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-186">x</span></span>||
|[<span data-ttu-id="bb4fa-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="bb4fa-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="bb4fa-188">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-188">x</span></span>|
|[<span data-ttu-id="bb4fa-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="bb4fa-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="bb4fa-190">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-190">x</span></span>|<span data-ttu-id="bb4fa-191">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-191">x</span></span>|<span data-ttu-id="bb4fa-192">x</span><span class="sxs-lookup"><span data-stu-id="bb4fa-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="bb4fa-193">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bb4fa-193">Attributes</span></span>

|<span data-ttu-id="bb4fa-194">Атрибут</span><span class="sxs-lookup"><span data-stu-id="bb4fa-194">Attribute</span></span>|<span data-ttu-id="bb4fa-195">Описание</span><span class="sxs-lookup"><span data-stu-id="bb4fa-195">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="bb4fa-196">xmlns</span><span class="sxs-lookup"><span data-stu-id="bb4fa-196">xmlns</span></span>|<span data-ttu-id="bb4fa-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="bb4fa-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="bb4fa-199">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="bb4fa-199">xmlns:xsi</span></span>|<span data-ttu-id="bb4fa-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="bb4fa-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="bb4fa-202">xsi:type</span><span class="sxs-lookup"><span data-stu-id="bb4fa-202">xsi:type</span></span>|<span data-ttu-id="bb4fa-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="bb4fa-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
