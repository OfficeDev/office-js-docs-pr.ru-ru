---
title: Элемент OfficeApp в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 42b6fe2e1c33322b90016d5e7ceec7b1bfe5b72d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433168"
---
# <a name="officeapp-element"></a><span data-ttu-id="c5d00-102">Элемент OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c5d00-102">OfficeApp element</span></span>

<span data-ttu-id="c5d00-103">Корневой элемент в манифесте надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="c5d00-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="c5d00-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="c5d00-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c5d00-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="c5d00-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="c5d00-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c5d00-106">Contained in</span></span>

 <span data-ttu-id="c5d00-107">_none_</span><span class="sxs-lookup"><span data-stu-id="c5d00-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="c5d00-108">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="c5d00-108">Must contain</span></span>

|<span data-ttu-id="c5d00-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="c5d00-109">**Element**</span></span>|<span data-ttu-id="c5d00-110">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="c5d00-110">**Content**</span></span>|<span data-ttu-id="c5d00-111">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="c5d00-111">**Mail**</span></span>|<span data-ttu-id="c5d00-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c5d00-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c5d00-113">Id</span><span class="sxs-lookup"><span data-stu-id="c5d00-113">Id</span></span>](id.md)|<span data-ttu-id="c5d00-114">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-114">x</span></span>|<span data-ttu-id="c5d00-115">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-115">x</span></span>|<span data-ttu-id="c5d00-116">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-116">x</span></span>|
|[<span data-ttu-id="c5d00-117">Version</span><span class="sxs-lookup"><span data-stu-id="c5d00-117">Version</span></span>](version.md)|<span data-ttu-id="c5d00-118">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-118">x</span></span>|<span data-ttu-id="c5d00-119">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-119">x</span></span>|<span data-ttu-id="c5d00-120">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-120">x</span></span>|
|[<span data-ttu-id="c5d00-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="c5d00-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="c5d00-122">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-122">x</span></span>|<span data-ttu-id="c5d00-123">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-123">x</span></span>|<span data-ttu-id="c5d00-124">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-124">x</span></span>|
|[<span data-ttu-id="c5d00-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="c5d00-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="c5d00-126">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-126">x</span></span>|<span data-ttu-id="c5d00-127">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-127">x</span></span>|<span data-ttu-id="c5d00-128">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-128">x</span></span>|
|[<span data-ttu-id="c5d00-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="c5d00-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="c5d00-130">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-130">x</span></span>||<span data-ttu-id="c5d00-131">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-131">x</span></span>|
|[<span data-ttu-id="c5d00-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="c5d00-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="c5d00-133">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-133">x</span></span>|<span data-ttu-id="c5d00-134">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-134">x</span></span>|<span data-ttu-id="c5d00-135">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-135">x</span></span>|
|[<span data-ttu-id="c5d00-136">Description</span><span class="sxs-lookup"><span data-stu-id="c5d00-136">Description</span></span>](description.md)|<span data-ttu-id="c5d00-137">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-137">x</span></span>|<span data-ttu-id="c5d00-138">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-138">x</span></span>|<span data-ttu-id="c5d00-139">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-139">x</span></span>|
|[<span data-ttu-id="c5d00-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="c5d00-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="c5d00-141">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-141">x</span></span>||
|[<span data-ttu-id="c5d00-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="c5d00-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="c5d00-143">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-143">x</span></span>||<span data-ttu-id="c5d00-144">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-144">x</span></span>|
|[<span data-ttu-id="c5d00-145">Rule</span><span class="sxs-lookup"><span data-stu-id="c5d00-145">Rule</span></span>](rule.md)||<span data-ttu-id="c5d00-146">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="c5d00-147">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="c5d00-147">Can contain</span></span>

|<span data-ttu-id="c5d00-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="c5d00-148">**Element**</span></span>|<span data-ttu-id="c5d00-149">**Контентная надстройка**</span><span class="sxs-lookup"><span data-stu-id="c5d00-149">**Content**</span></span>|<span data-ttu-id="c5d00-150">**Почтовая надстройка**</span><span class="sxs-lookup"><span data-stu-id="c5d00-150">**Mail**</span></span>|<span data-ttu-id="c5d00-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="c5d00-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c5d00-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="c5d00-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="c5d00-153">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-153">x</span></span>|<span data-ttu-id="c5d00-154">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-154">x</span></span>|<span data-ttu-id="c5d00-155">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-155">x</span></span>|
|[<span data-ttu-id="c5d00-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="c5d00-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="c5d00-157">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-157">x</span></span>|<span data-ttu-id="c5d00-158">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-158">x</span></span>|<span data-ttu-id="c5d00-159">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-159">x</span></span>|
|[<span data-ttu-id="c5d00-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="c5d00-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="c5d00-161">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-161">x</span></span>|<span data-ttu-id="c5d00-162">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-162">x</span></span>|<span data-ttu-id="c5d00-163">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-163">x</span></span>|
|[<span data-ttu-id="c5d00-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="c5d00-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="c5d00-165">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-165">x</span></span>|<span data-ttu-id="c5d00-166">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-166">x</span></span>|<span data-ttu-id="c5d00-167">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-167">x</span></span>|
|[<span data-ttu-id="c5d00-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="c5d00-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="c5d00-169">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-169">x</span></span>|<span data-ttu-id="c5d00-170">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-170">x</span></span>|<span data-ttu-id="c5d00-171">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-171">x</span></span>|
|[<span data-ttu-id="c5d00-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="c5d00-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="c5d00-173">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-173">x</span></span>|<span data-ttu-id="c5d00-174">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-174">x</span></span>|<span data-ttu-id="c5d00-175">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-175">x</span></span>|
|[<span data-ttu-id="c5d00-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5d00-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="c5d00-177">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-177">x</span></span>|<span data-ttu-id="c5d00-178">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-178">x</span></span>|<span data-ttu-id="c5d00-179">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-179">x</span></span>|
|[<span data-ttu-id="c5d00-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="c5d00-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="c5d00-181">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-181">x</span></span>|||
|[<span data-ttu-id="c5d00-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="c5d00-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="c5d00-183">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-183">x</span></span>||
|[<span data-ttu-id="c5d00-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="c5d00-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="c5d00-185">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-185">x</span></span>||
|[<span data-ttu-id="c5d00-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="c5d00-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="c5d00-187">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-187">x</span></span>|
|[<span data-ttu-id="c5d00-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="c5d00-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="c5d00-189">x</span><span class="sxs-lookup"><span data-stu-id="c5d00-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="c5d00-190">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c5d00-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="c5d00-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="c5d00-191">xmlns</span></span>|<span data-ttu-id="c5d00-p101">Определяет пространство имен и версию схемы для манифеста надстройки Office. Для этого атрибута всегда должно быть задано значение `"http://schemas.microsoft.com/office/appforoffice/1.1"`.</span><span class="sxs-lookup"><span data-stu-id="c5d00-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="c5d00-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="c5d00-194">xmlns:xsi</span></span>|<span data-ttu-id="c5d00-p102">Определяет экземпляр объекта XMLSchema. Для этого атрибута всегда должно быть задано значение `"http://www.w3.org/2001/XMLSchema-instance"`.</span><span class="sxs-lookup"><span data-stu-id="c5d00-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="c5d00-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="c5d00-197">xsi:type</span></span>|<span data-ttu-id="c5d00-p103">Определяет тип надстройки Office. Для этого атрибута должно быть задано одно из следующих значений: `"ContentApp"`, `"MailApp"` или `"TaskPaneApp"`</span><span class="sxs-lookup"><span data-stu-id="c5d00-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
