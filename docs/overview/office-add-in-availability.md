---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 36c6bc6b6348ac988049f9a50127f6dd2f94bf37
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217825"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9c31a-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9c31a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9c31a-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="9c31a-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="9c31a-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="9c31a-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="9c31a-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="9c31a-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="9c31a-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="9c31a-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="9c31a-108">Excel</span><span class="sxs-lookup"><span data-stu-id="9c31a-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9c31a-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9c31a-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9c31a-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9c31a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="9c31a-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-114">- TaskPane</span></span><br><span data-ttu-id="9c31a-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-115">
        - Content</span></span><br><span data-ttu-id="9c31a-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9c31a-116">
        - Custom Functions</span></span><br><span data-ttu-id="9c31a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="9c31a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9c31a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9c31a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9c31a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9c31a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9c31a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9c31a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9c31a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9c31a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9c31a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9c31a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="9c31a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="9c31a-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-131">
        - BindingEvents</span></span><br><span data-ttu-id="9c31a-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-132">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-133">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-134">
        - File</span></span><br><span data-ttu-id="9c31a-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-135">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-137">
        - Selection</span></span><br><span data-ttu-id="9c31a-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-138">
        - Settings</span></span><br><span data-ttu-id="9c31a-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-139">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-140">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-141">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-143">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-143">Office on Windows</span></span><br><span data-ttu-id="9c31a-144">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-145">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-145">- TaskPane</span></span><br><span data-ttu-id="9c31a-146">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-146">
        - Content</span></span><br><span data-ttu-id="9c31a-147">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9c31a-147">
        - Custom Functions</span></span><br><span data-ttu-id="9c31a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="9c31a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9c31a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9c31a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9c31a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9c31a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9c31a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9c31a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9c31a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9c31a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9c31a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9c31a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="9c31a-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="9c31a-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-163">
        - BindingEvents</span></span><br><span data-ttu-id="9c31a-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-164">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-165">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-166">
        - File</span></span><br><span data-ttu-id="9c31a-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-167">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-169">
        - Selection</span></span><br><span data-ttu-id="9c31a-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-170">
        - Settings</span></span><br><span data-ttu-id="9c31a-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-171">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-172">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-173">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-175">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-175">Office 2019 on Windows</span></span><br><span data-ttu-id="9c31a-176">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9c31a-177">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-177">- TaskPane</span></span><br><span data-ttu-id="9c31a-178">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-178">
        - Content</span></span><br><span data-ttu-id="9c31a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9c31a-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9c31a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9c31a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9c31a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9c31a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9c31a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9c31a-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9c31a-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-190">- BindingEvents</span></span><br><span data-ttu-id="9c31a-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-191">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-192">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-193">
        - File</span></span><br><span data-ttu-id="9c31a-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-194">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-196">
        - Selection</span></span><br><span data-ttu-id="9c31a-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-197">
        - Settings</span></span><br><span data-ttu-id="9c31a-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-198">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-199">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-200">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-202">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-202">Office 2016 on Windows</span></span><br><span data-ttu-id="9c31a-203">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9c31a-204">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-204">- TaskPane</span></span><br><span data-ttu-id="9c31a-205">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-205">
        - Content</span></span></td>
    <td><span data-ttu-id="9c31a-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9c31a-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-209">- BindingEvents</span></span><br><span data-ttu-id="9c31a-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-210">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-211">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-212">
        - File</span></span><br><span data-ttu-id="9c31a-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-213">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-215">
        - Selection</span></span><br><span data-ttu-id="9c31a-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-216">
        - Settings</span></span><br><span data-ttu-id="9c31a-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-217">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-218">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-219">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-221">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-221">Office 2013 on Windows</span></span><br><span data-ttu-id="9c31a-222">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9c31a-223">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-223">
        - TaskPane</span></span><br><span data-ttu-id="9c31a-224">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9c31a-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9c31a-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9c31a-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-227">
        - BindingEvents</span></span><br><span data-ttu-id="9c31a-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-228">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-229">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-230">
        - File</span></span><br><span data-ttu-id="9c31a-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-231">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-233">
        - Selection</span></span><br><span data-ttu-id="9c31a-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-234">
        - Settings</span></span><br><span data-ttu-id="9c31a-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-235">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-236">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-237">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-239">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9c31a-239">Office on iPad</span></span><br><span data-ttu-id="9c31a-240">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="9c31a-241">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-241">- TaskPane</span></span><br><span data-ttu-id="9c31a-242">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-242">
        - Content</span></span></td>
    <td><span data-ttu-id="9c31a-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9c31a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9c31a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9c31a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9c31a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9c31a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9c31a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9c31a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9c31a-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9c31a-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="9c31a-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-256">- BindingEvents</span></span><br><span data-ttu-id="9c31a-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-257">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-258">
        - File</span></span><br><span data-ttu-id="9c31a-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-259">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-261">
        - Selection</span></span><br><span data-ttu-id="9c31a-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-262">
        - Settings</span></span><br><span data-ttu-id="9c31a-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-263">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-264">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-265">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-267">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-267">Office on Mac</span></span><br><span data-ttu-id="9c31a-268">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="9c31a-269">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-269">- TaskPane</span></span><br><span data-ttu-id="9c31a-270">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-270">
        - Content</span></span><br><span data-ttu-id="9c31a-271">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9c31a-271">
        - Custom Functions</span></span><br><span data-ttu-id="9c31a-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9c31a-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9c31a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9c31a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9c31a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9c31a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9c31a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9c31a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9c31a-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="9c31a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="9c31a-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="9c31a-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="9c31a-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-287">- BindingEvents</span></span><br><span data-ttu-id="9c31a-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-288">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-289">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-290">
        - File</span></span><br><span data-ttu-id="9c31a-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-291">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-293">
        - PdfFile</span></span><br><span data-ttu-id="9c31a-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-294">
        - Selection</span></span><br><span data-ttu-id="9c31a-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-295">
        - Settings</span></span><br><span data-ttu-id="9c31a-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-296">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-297">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-298">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-300">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-300">Office 2019 on Mac</span></span><br><span data-ttu-id="9c31a-301">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9c31a-302">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-302">- TaskPane</span></span><br><span data-ttu-id="9c31a-303">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-303">
        - Content</span></span><br><span data-ttu-id="9c31a-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9c31a-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9c31a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9c31a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9c31a-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9c31a-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9c31a-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9c31a-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9c31a-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-315">- BindingEvents</span></span><br><span data-ttu-id="9c31a-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-316">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-317">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-318">
        - File</span></span><br><span data-ttu-id="9c31a-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-319">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-321">
        - PdfFile</span></span><br><span data-ttu-id="9c31a-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-322">
        - Selection</span></span><br><span data-ttu-id="9c31a-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-323">
        - Settings</span></span><br><span data-ttu-id="9c31a-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-324">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-325">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-326">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-328">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-328">Office 2016 on Mac</span></span><br><span data-ttu-id="9c31a-329">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="9c31a-330">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-330">- TaskPane</span></span><br><span data-ttu-id="9c31a-331">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-331">
        - Content</span></span></td>
    <td><span data-ttu-id="9c31a-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9c31a-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9c31a-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="9c31a-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-335">- BindingEvents</span></span><br><span data-ttu-id="9c31a-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-336">
        - CompressedFile</span></span><br><span data-ttu-id="9c31a-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-337">
        - DocumentEvents</span></span><br><span data-ttu-id="9c31a-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-338">
        - File</span></span><br><span data-ttu-id="9c31a-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-339">
        - MatrixBindings</span></span><br><span data-ttu-id="9c31a-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-341">
        - PdfFile</span></span><br><span data-ttu-id="9c31a-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-342">
        - Selection</span></span><br><span data-ttu-id="9c31a-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-343">
        - Settings</span></span><br><span data-ttu-id="9c31a-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-344">
        - TableBindings</span></span><br><span data-ttu-id="9c31a-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-345">
        - TableCoercion</span></span><br><span data-ttu-id="9c31a-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-346">
        - TextBindings</span></span><br><span data-ttu-id="9c31a-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="9c31a-348">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9c31a-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="9c31a-349">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="9c31a-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9c31a-350">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9c31a-351">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9c31a-352">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9c31a-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-354">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-354">Office on the web</span></span></td>
    <td><span data-ttu-id="9c31a-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9c31a-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="9c31a-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-357">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-357">Office on Windows</span></span><br><span data-ttu-id="9c31a-358">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="9c31a-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9c31a-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="9c31a-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-361">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-361">Office for Mac</span></span><br><span data-ttu-id="9c31a-362">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="9c31a-363">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="9c31a-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="9c31a-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="9c31a-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="9c31a-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9c31a-366">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-366">Platform</span></span></th>
    <th><span data-ttu-id="9c31a-367">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-367">Extension points</span></span></th>
    <th><span data-ttu-id="9c31a-368">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="9c31a-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-370">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-370">Office on the web</span></span><br><span data-ttu-id="9c31a-371">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="9c31a-371">(modern)</span></span></td>
    <td> <span data-ttu-id="9c31a-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9c31a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="9c31a-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="9c31a-385">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-386">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-386">Office on the web</span></span><br><span data-ttu-id="9c31a-387">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="9c31a-387">(classic)</span></span></td>
    <td> <span data-ttu-id="9c31a-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9c31a-399">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-400">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-400">Office on Windows</span></span><br><span data-ttu-id="9c31a-401">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9c31a-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="9c31a-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9c31a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="9c31a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="9c31a-416">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-417">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-417">Office 2019 on Windows</span></span><br><span data-ttu-id="9c31a-418">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9c31a-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="9c31a-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9c31a-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9c31a-432">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-433">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-433">Office 2016 on Windows</span></span><br><span data-ttu-id="9c31a-434">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9c31a-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="9c31a-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="9c31a-445">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-446">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-446">Office 2013 on Windows</span></span><br><span data-ttu-id="9c31a-447">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="9c31a-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="9c31a-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="9c31a-456">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-457">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="9c31a-457">Office on iOS</span></span><br><span data-ttu-id="9c31a-458">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9c31a-466">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-467">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-467">Office on Mac</span></span><br><span data-ttu-id="9c31a-468">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9c31a-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="9c31a-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="9c31a-482">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-483">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-483">Office 2019 on Mac</span></span><br><span data-ttu-id="9c31a-484">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9c31a-496">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-497">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-497">Office 2016 on Mac</span></span><br><span data-ttu-id="9c31a-498">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="9c31a-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="9c31a-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="9c31a-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9c31a-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9c31a-510">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-511">Office для Android</span><span class="sxs-lookup"><span data-stu-id="9c31a-511">Office on Android</span></span><br><span data-ttu-id="9c31a-512">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="9c31a-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Организатор встречи (создание): собрание по сети</a> (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9c31a-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="9c31a-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9c31a-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9c31a-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9c31a-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9c31a-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9c31a-521">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9c31a-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="9c31a-522">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9c31a-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9c31a-523">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="9c31a-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="9c31a-524">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="9c31a-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="9c31a-525">Word</span><span class="sxs-lookup"><span data-stu-id="9c31a-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9c31a-526">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-526">Platform</span></span></th>
    <th><span data-ttu-id="9c31a-527">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-527">Extension points</span></span></th>
    <th><span data-ttu-id="9c31a-528">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="9c31a-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-530">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="9c31a-531">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-531">- TaskPane</span></span><br><span data-ttu-id="9c31a-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9c31a-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9c31a-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9c31a-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-539">- BindingEvents</span></span><br><span data-ttu-id="9c31a-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-541">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-542">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-542">
         - File</span></span><br><span data-ttu-id="9c31a-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-544">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-547">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-548">
         - Selection</span></span><br><span data-ttu-id="9c31a-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-549">
         - Settings</span></span><br><span data-ttu-id="9c31a-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-550">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-551">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-552">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-553">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-555">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-555">Office on Windows</span></span><br><span data-ttu-id="9c31a-556">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-557">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-557">- TaskPane</span></span><br><span data-ttu-id="9c31a-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9c31a-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9c31a-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9c31a-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-565">- BindingEvents</span></span><br><span data-ttu-id="9c31a-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-566">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-568">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-569">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-569">
         - File</span></span><br><span data-ttu-id="9c31a-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-571">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-574">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-575">
         - Selection</span></span><br><span data-ttu-id="9c31a-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-576">
         - Settings</span></span><br><span data-ttu-id="9c31a-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-577">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-578">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-579">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-580">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-582">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-582">Office 2019 on Windows</span></span><br><span data-ttu-id="9c31a-583">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-584">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-584">- TaskPane</span></span><br><span data-ttu-id="9c31a-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9c31a-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9c31a-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-591">- BindingEvents</span></span><br><span data-ttu-id="9c31a-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-592">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-594">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-595">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-595">
         - File</span></span><br><span data-ttu-id="9c31a-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-597">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-600">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-601">
         - Selection</span></span><br><span data-ttu-id="9c31a-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-602">
         - Settings</span></span><br><span data-ttu-id="9c31a-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-603">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-604">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-605">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-606">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-608">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-608">Office 2016 on Windows</span></span><br><span data-ttu-id="9c31a-609">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-610">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9c31a-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-614">- BindingEvents</span></span><br><span data-ttu-id="9c31a-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-615">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-617">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-618">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-618">
         - File</span></span><br><span data-ttu-id="9c31a-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-620">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-623">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-624">
         - Selection</span></span><br><span data-ttu-id="9c31a-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-625">
         - Settings</span></span><br><span data-ttu-id="9c31a-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-626">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-627">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-628">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-629">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-631">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-631">Office 2013 on Windows</span></span><br><span data-ttu-id="9c31a-632">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-633">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9c31a-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9c31a-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-636">- BindingEvents</span></span><br><span data-ttu-id="9c31a-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-637">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-639">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-640">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-640">
         - File</span></span><br><span data-ttu-id="9c31a-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-642">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-645">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-646">
         - Selection</span></span><br><span data-ttu-id="9c31a-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-647">
         - Settings</span></span><br><span data-ttu-id="9c31a-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-648">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-649">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-650">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-651">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-653">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9c31a-653">Office on iPad</span></span><br><span data-ttu-id="9c31a-654">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-655">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9c31a-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9c31a-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="9c31a-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-661">- BindingEvents</span></span><br><span data-ttu-id="9c31a-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-662">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-664">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-665">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-665">
         - File</span></span><br><span data-ttu-id="9c31a-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-667">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-670">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-671">
         - Selection</span></span><br><span data-ttu-id="9c31a-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-672">
         - Settings</span></span><br><span data-ttu-id="9c31a-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-673">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-674">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-675">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-676">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-678">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-678">Office on Mac</span></span><br><span data-ttu-id="9c31a-679">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-680">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-680">- TaskPane</span></span><br><span data-ttu-id="9c31a-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9c31a-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9c31a-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="9c31a-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-688">- BindingEvents</span></span><br><span data-ttu-id="9c31a-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-689">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-691">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-692">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-692">
         - File</span></span><br><span data-ttu-id="9c31a-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-694">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-697">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-698">
         - Selection</span></span><br><span data-ttu-id="9c31a-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-699">
         - Settings</span></span><br><span data-ttu-id="9c31a-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-700">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-701">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-702">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-703">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-705">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-705">Office 2019 on Mac</span></span><br><span data-ttu-id="9c31a-706">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-707">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-707">- TaskPane</span></span><br><span data-ttu-id="9c31a-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="9c31a-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="9c31a-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="9c31a-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-714">- BindingEvents</span></span><br><span data-ttu-id="9c31a-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-715">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-717">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-718">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-718">
         - File</span></span><br><span data-ttu-id="9c31a-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-720">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-723">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-724">
         - Selection</span></span><br><span data-ttu-id="9c31a-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-725">
         - Settings</span></span><br><span data-ttu-id="9c31a-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-726">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-727">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-728">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-729">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-731">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-731">Office 2016 on Mac</span></span><br><span data-ttu-id="9c31a-732">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-733">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="9c31a-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9c31a-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="9c31a-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-737">- BindingEvents</span></span><br><span data-ttu-id="9c31a-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-738">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9c31a-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="9c31a-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-740">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-741">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9c31a-741">
         - File</span></span><br><span data-ttu-id="9c31a-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-743">
         - MatrixBindings</span></span><br><span data-ttu-id="9c31a-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="9c31a-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9c31a-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-746">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-747">
         - Selection</span></span><br><span data-ttu-id="9c31a-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9c31a-748">
         - Settings</span></span><br><span data-ttu-id="9c31a-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-749">
         - TableBindings</span></span><br><span data-ttu-id="9c31a-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-750">
         - TableCoercion</span></span><br><span data-ttu-id="9c31a-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9c31a-751">
         - TextBindings</span></span><br><span data-ttu-id="9c31a-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-752">
         - TextCoercion</span></span><br><span data-ttu-id="9c31a-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="9c31a-754">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9c31a-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9c31a-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9c31a-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9c31a-756">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-756">Platform</span></span></th>
    <th><span data-ttu-id="9c31a-757">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-757">Extension points</span></span></th>
    <th><span data-ttu-id="9c31a-758">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="9c31a-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-760">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="9c31a-761">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-761">- Content</span></span><br><span data-ttu-id="9c31a-762">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-762">
         - TaskPane</span></span><br><span data-ttu-id="9c31a-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9c31a-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9c31a-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-768">- ActiveView</span></span><br><span data-ttu-id="9c31a-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-769">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-770">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-771">
         - File</span></span><br><span data-ttu-id="9c31a-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-772">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-773">
         - Selection</span></span><br><span data-ttu-id="9c31a-774">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-774">
         - Settings</span></span><br><span data-ttu-id="9c31a-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-776">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-776">Office on Windows</span></span><br><span data-ttu-id="9c31a-777">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-778">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-778">- Content</span></span><br><span data-ttu-id="9c31a-779">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-779">
         - TaskPane</span></span><br><span data-ttu-id="9c31a-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9c31a-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9c31a-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-785">- ActiveView</span></span><br><span data-ttu-id="9c31a-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-786">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-787">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-788">
         - File</span></span><br><span data-ttu-id="9c31a-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-789">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-790">
         - Selection</span></span><br><span data-ttu-id="9c31a-791">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-791">
         - Settings</span></span><br><span data-ttu-id="9c31a-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-793">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-793">Office 2019 on Windows</span></span><br><span data-ttu-id="9c31a-794">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-795">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-795">- Content</span></span><br><span data-ttu-id="9c31a-796">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-796">
         - TaskPane</span></span><br><span data-ttu-id="9c31a-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-800">- ActiveView</span></span><br><span data-ttu-id="9c31a-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-801">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-802">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-803">
         - File</span></span><br><span data-ttu-id="9c31a-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-804">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-805">
         - Selection</span></span><br><span data-ttu-id="9c31a-806">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-806">
         - Settings</span></span><br><span data-ttu-id="9c31a-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-808">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-808">Office 2016 on Windows</span></span><br><span data-ttu-id="9c31a-809">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-810">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-810">- Content</span></span><br><span data-ttu-id="9c31a-811">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9c31a-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9c31a-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-814">- ActiveView</span></span><br><span data-ttu-id="9c31a-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-815">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-816">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-817">
         - File</span></span><br><span data-ttu-id="9c31a-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-818">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-819">
         - Selection</span></span><br><span data-ttu-id="9c31a-820">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-820">
         - Settings</span></span><br><span data-ttu-id="9c31a-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-822">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-822">Office 2013 on Windows</span></span><br><span data-ttu-id="9c31a-823">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-824">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-824">- Content</span></span><br><span data-ttu-id="9c31a-825">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="9c31a-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9c31a-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9c31a-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-828">- ActiveView</span></span><br><span data-ttu-id="9c31a-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-829">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-830">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-831">
         - File</span></span><br><span data-ttu-id="9c31a-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-832">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-833">
         - Selection</span></span><br><span data-ttu-id="9c31a-834">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-834">
         - Settings</span></span><br><span data-ttu-id="9c31a-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-836">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9c31a-836">Office on iPad</span></span><br><span data-ttu-id="9c31a-837">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-838">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-838">- Content</span></span><br><span data-ttu-id="9c31a-839">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9c31a-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-843">- ActiveView</span></span><br><span data-ttu-id="9c31a-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-844">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-845">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-846">
         - File</span></span><br><span data-ttu-id="9c31a-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-847">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-848">
         - Selection</span></span><br><span data-ttu-id="9c31a-849">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-849">
         - Settings</span></span><br><span data-ttu-id="9c31a-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-851">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-851">Office on Mac</span></span><br><span data-ttu-id="9c31a-852">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c31a-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="9c31a-853">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-853">- Content</span></span><br><span data-ttu-id="9c31a-854">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-854">
         - TaskPane</span></span><br><span data-ttu-id="9c31a-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="9c31a-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="9c31a-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="9c31a-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-860">- ActiveView</span></span><br><span data-ttu-id="9c31a-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-861">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-862">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-863">
         - File</span></span><br><span data-ttu-id="9c31a-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-864">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-865">
         - Selection</span></span><br><span data-ttu-id="9c31a-866">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-866">
         - Settings</span></span><br><span data-ttu-id="9c31a-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-868">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-868">Office 2019 on Mac</span></span><br><span data-ttu-id="9c31a-869">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-870">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-870">- Content</span></span><br><span data-ttu-id="9c31a-871">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-871">
         - TaskPane</span></span><br><span data-ttu-id="9c31a-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-875">- ActiveView</span></span><br><span data-ttu-id="9c31a-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-876">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-877">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-878">
         - File</span></span><br><span data-ttu-id="9c31a-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-879">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-880">
         - Selection</span></span><br><span data-ttu-id="9c31a-881">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-881">
         - Settings</span></span><br><span data-ttu-id="9c31a-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-883">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-883">Office 2016 on Mac</span></span><br><span data-ttu-id="9c31a-884">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-885">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-885">- Content</span></span><br><span data-ttu-id="9c31a-886">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9c31a-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="9c31a-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9c31a-889">- ActiveView</span></span><br><span data-ttu-id="9c31a-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-890">
         - CompressedFile</span></span><br><span data-ttu-id="9c31a-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-891">
         - DocumentEvents</span></span><br><span data-ttu-id="9c31a-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="9c31a-892">
         - File</span></span><br><span data-ttu-id="9c31a-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9c31a-893">
         - PdfFile</span></span><br><span data-ttu-id="9c31a-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-894">
         - Selection</span></span><br><span data-ttu-id="9c31a-895">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-895">
         - Settings</span></span><br><span data-ttu-id="9c31a-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="9c31a-897">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="9c31a-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="9c31a-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="9c31a-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9c31a-899">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-899">Platform</span></span></th>
    <th><span data-ttu-id="9c31a-900">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-900">Extension points</span></span></th>
    <th><span data-ttu-id="9c31a-901">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="9c31a-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-903">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="9c31a-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="9c31a-904">- Контент</span><span class="sxs-lookup"><span data-stu-id="9c31a-904">- Content</span></span><br><span data-ttu-id="9c31a-905">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-905">
         - TaskPane</span></span><br><span data-ttu-id="9c31a-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9c31a-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9c31a-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="9c31a-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9c31a-910">- DocumentEvents</span></span><br><span data-ttu-id="9c31a-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="9c31a-912">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9c31a-912">
         - Settings</span></span><br><span data-ttu-id="9c31a-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="9c31a-914">Project</span><span class="sxs-lookup"><span data-stu-id="9c31a-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9c31a-915">Платформа</span><span class="sxs-lookup"><span data-stu-id="9c31a-915">Platform</span></span></th>
    <th><span data-ttu-id="9c31a-916">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9c31a-916">Extension points</span></span></th>
    <th><span data-ttu-id="9c31a-917">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9c31a-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="9c31a-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9c31a-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-919">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-919">Office 2019 on Windows</span></span><br><span data-ttu-id="9c31a-920">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-921">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-923">- Selection</span></span><br><span data-ttu-id="9c31a-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-925">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-925">Office 2016 on Windows</span></span><br><span data-ttu-id="9c31a-926">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-927">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-929">- Selection</span></span><br><span data-ttu-id="9c31a-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9c31a-931">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9c31a-931">Office 2013 on Windows</span></span><br><span data-ttu-id="9c31a-932">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="9c31a-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="9c31a-933">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9c31a-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9c31a-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9c31a-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9c31a-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="9c31a-935">- Selection</span></span><br><span data-ttu-id="9c31a-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9c31a-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9c31a-937">См. также</span><span class="sxs-lookup"><span data-stu-id="9c31a-937">See also</span></span>

- [<span data-ttu-id="9c31a-938">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9c31a-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9c31a-939">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="9c31a-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="9c31a-940">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="9c31a-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="9c31a-941">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="9c31a-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="9c31a-942">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="9c31a-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="9c31a-943">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="9c31a-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="9c31a-944">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="9c31a-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="9c31a-945">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="9c31a-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="9c31a-946">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="9c31a-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="9c31a-947">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="9c31a-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="9c31a-948">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="9c31a-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="9c31a-949">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9c31a-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)