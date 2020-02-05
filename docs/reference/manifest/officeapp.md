---
title: Элемент OfficeApp в файле манифеста
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 080025e62a56421dff942792f99ee672ce1db69a
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773582"
---
# <a name="officeapp-element"></a><span data-ttu-id="11a03-102">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="11a03-102">OfficeApp element</span></span>

<span data-ttu-id="11a03-103">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="11a03-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="11a03-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="11a03-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="11a03-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="11a03-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="11a03-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="11a03-106">Contained in</span></span>

 <span data-ttu-id="11a03-107">_none_</span><span class="sxs-lookup"><span data-stu-id="11a03-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="11a03-108">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="11a03-108">Must contain</span></span>

|<span data-ttu-id="11a03-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="11a03-109">**Element**</span></span>|<span data-ttu-id="11a03-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="11a03-110">**Content**</span></span>|<span data-ttu-id="11a03-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="11a03-111">**Mail**</span></span>|<span data-ttu-id="11a03-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="11a03-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="11a03-113">Id</span><span class="sxs-lookup"><span data-stu-id="11a03-113">Id</span></span>](id.md)|<span data-ttu-id="11a03-114">x</span><span class="sxs-lookup"><span data-stu-id="11a03-114">x</span></span>|<span data-ttu-id="11a03-115">x</span><span class="sxs-lookup"><span data-stu-id="11a03-115">x</span></span>|<span data-ttu-id="11a03-116">x</span><span class="sxs-lookup"><span data-stu-id="11a03-116">x</span></span>|
|[<span data-ttu-id="11a03-117">Версия</span><span class="sxs-lookup"><span data-stu-id="11a03-117">Version</span></span>](version.md)|<span data-ttu-id="11a03-118">x</span><span class="sxs-lookup"><span data-stu-id="11a03-118">x</span></span>|<span data-ttu-id="11a03-119">x</span><span class="sxs-lookup"><span data-stu-id="11a03-119">x</span></span>|<span data-ttu-id="11a03-120">x</span><span class="sxs-lookup"><span data-stu-id="11a03-120">x</span></span>|
|[<span data-ttu-id="11a03-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="11a03-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="11a03-122">x</span><span class="sxs-lookup"><span data-stu-id="11a03-122">x</span></span>|<span data-ttu-id="11a03-123">x</span><span class="sxs-lookup"><span data-stu-id="11a03-123">x</span></span>|<span data-ttu-id="11a03-124">x</span><span class="sxs-lookup"><span data-stu-id="11a03-124">x</span></span>|
|[<span data-ttu-id="11a03-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="11a03-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="11a03-126">x</span><span class="sxs-lookup"><span data-stu-id="11a03-126">x</span></span>|<span data-ttu-id="11a03-127">x</span><span class="sxs-lookup"><span data-stu-id="11a03-127">x</span></span>|<span data-ttu-id="11a03-128">x</span><span class="sxs-lookup"><span data-stu-id="11a03-128">x</span></span>|
|[<span data-ttu-id="11a03-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="11a03-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="11a03-130">x</span><span class="sxs-lookup"><span data-stu-id="11a03-130">x</span></span>||<span data-ttu-id="11a03-131">x</span><span class="sxs-lookup"><span data-stu-id="11a03-131">x</span></span>|
|[<span data-ttu-id="11a03-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="11a03-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="11a03-133">x</span><span class="sxs-lookup"><span data-stu-id="11a03-133">x</span></span>|<span data-ttu-id="11a03-134">x</span><span class="sxs-lookup"><span data-stu-id="11a03-134">x</span></span>|<span data-ttu-id="11a03-135">x</span><span class="sxs-lookup"><span data-stu-id="11a03-135">x</span></span>|
|[<span data-ttu-id="11a03-136">Описание</span><span class="sxs-lookup"><span data-stu-id="11a03-136">Description</span></span>](description.md)|<span data-ttu-id="11a03-137">x</span><span class="sxs-lookup"><span data-stu-id="11a03-137">x</span></span>|<span data-ttu-id="11a03-138">x</span><span class="sxs-lookup"><span data-stu-id="11a03-138">x</span></span>|<span data-ttu-id="11a03-139">x</span><span class="sxs-lookup"><span data-stu-id="11a03-139">x</span></span>|
|[<span data-ttu-id="11a03-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="11a03-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="11a03-141">x</span><span class="sxs-lookup"><span data-stu-id="11a03-141">x</span></span>||
|[<span data-ttu-id="11a03-142">Разрешения</span><span class="sxs-lookup"><span data-stu-id="11a03-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="11a03-143">x</span><span class="sxs-lookup"><span data-stu-id="11a03-143">x</span></span>||<span data-ttu-id="11a03-144">x</span><span class="sxs-lookup"><span data-stu-id="11a03-144">x</span></span>|
|[<span data-ttu-id="11a03-145">Rule</span><span class="sxs-lookup"><span data-stu-id="11a03-145">Rule</span></span>](rule.md)||<span data-ttu-id="11a03-146">x</span><span class="sxs-lookup"><span data-stu-id="11a03-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="11a03-147">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="11a03-147">Can contain</span></span>

|<span data-ttu-id="11a03-148">**Элемент**</span><span class="sxs-lookup"><span data-stu-id="11a03-148">**Element**</span></span>|<span data-ttu-id="11a03-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="11a03-149">**Content**</span></span>|<span data-ttu-id="11a03-150">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="11a03-150">**Mail**</span></span>|<span data-ttu-id="11a03-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="11a03-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="11a03-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="11a03-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="11a03-153">x</span><span class="sxs-lookup"><span data-stu-id="11a03-153">x</span></span>|<span data-ttu-id="11a03-154">x</span><span class="sxs-lookup"><span data-stu-id="11a03-154">x</span></span>|<span data-ttu-id="11a03-155">x</span><span class="sxs-lookup"><span data-stu-id="11a03-155">x</span></span>|
|[<span data-ttu-id="11a03-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="11a03-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="11a03-157">x</span><span class="sxs-lookup"><span data-stu-id="11a03-157">x</span></span>|<span data-ttu-id="11a03-158">x</span><span class="sxs-lookup"><span data-stu-id="11a03-158">x</span></span>|<span data-ttu-id="11a03-159">x</span><span class="sxs-lookup"><span data-stu-id="11a03-159">x</span></span>|
|[<span data-ttu-id="11a03-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="11a03-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="11a03-161">x</span><span class="sxs-lookup"><span data-stu-id="11a03-161">x</span></span>|<span data-ttu-id="11a03-162">x</span><span class="sxs-lookup"><span data-stu-id="11a03-162">x</span></span>|<span data-ttu-id="11a03-163">x</span><span class="sxs-lookup"><span data-stu-id="11a03-163">x</span></span>|
|[<span data-ttu-id="11a03-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="11a03-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="11a03-165">x</span><span class="sxs-lookup"><span data-stu-id="11a03-165">x</span></span>|<span data-ttu-id="11a03-166">x</span><span class="sxs-lookup"><span data-stu-id="11a03-166">x</span></span>|<span data-ttu-id="11a03-167">x</span><span class="sxs-lookup"><span data-stu-id="11a03-167">x</span></span>|
|[<span data-ttu-id="11a03-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="11a03-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="11a03-169">x</span><span class="sxs-lookup"><span data-stu-id="11a03-169">x</span></span>|<span data-ttu-id="11a03-170">x</span><span class="sxs-lookup"><span data-stu-id="11a03-170">x</span></span>|<span data-ttu-id="11a03-171">x</span><span class="sxs-lookup"><span data-stu-id="11a03-171">x</span></span>|
|[<span data-ttu-id="11a03-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="11a03-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="11a03-173">x</span><span class="sxs-lookup"><span data-stu-id="11a03-173">x</span></span>|<span data-ttu-id="11a03-174">x</span><span class="sxs-lookup"><span data-stu-id="11a03-174">x</span></span>|<span data-ttu-id="11a03-175">x</span><span class="sxs-lookup"><span data-stu-id="11a03-175">x</span></span>|
|[<span data-ttu-id="11a03-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="11a03-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="11a03-177">x</span><span class="sxs-lookup"><span data-stu-id="11a03-177">x</span></span>|<span data-ttu-id="11a03-178">x</span><span class="sxs-lookup"><span data-stu-id="11a03-178">x</span></span>|<span data-ttu-id="11a03-179">x</span><span class="sxs-lookup"><span data-stu-id="11a03-179">x</span></span>|
|[<span data-ttu-id="11a03-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="11a03-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="11a03-181">x</span><span class="sxs-lookup"><span data-stu-id="11a03-181">x</span></span>|||
|[<span data-ttu-id="11a03-182">Разрешения</span><span class="sxs-lookup"><span data-stu-id="11a03-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="11a03-183">x</span><span class="sxs-lookup"><span data-stu-id="11a03-183">x</span></span>||
|[<span data-ttu-id="11a03-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="11a03-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="11a03-185">x</span><span class="sxs-lookup"><span data-stu-id="11a03-185">x</span></span>||
|[<span data-ttu-id="11a03-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="11a03-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="11a03-187">x</span><span class="sxs-lookup"><span data-stu-id="11a03-187">x</span></span>|
|[<span data-ttu-id="11a03-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="11a03-188">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="11a03-189">x</span><span class="sxs-lookup"><span data-stu-id="11a03-189">x</span></span>|<span data-ttu-id="11a03-190">x</span><span class="sxs-lookup"><span data-stu-id="11a03-190">x</span></span>|<span data-ttu-id="11a03-191">x</span><span class="sxs-lookup"><span data-stu-id="11a03-191">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="11a03-192">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="11a03-192">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="11a03-193">xmlns</span><span class="sxs-lookup"><span data-stu-id="11a03-193">xmlns</span></span>|<span data-ttu-id="11a03-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="11a03-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="11a03-196">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="11a03-196">xmlns:xsi</span></span>|<span data-ttu-id="11a03-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="11a03-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="11a03-199">xsi:type</span><span class="sxs-lookup"><span data-stu-id="11a03-199">xsi:type</span></span>|<span data-ttu-id="11a03-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="11a03-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
