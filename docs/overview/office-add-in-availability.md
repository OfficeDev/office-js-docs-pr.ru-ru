---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 05/23/2019
localization_priority: Priority
ms.openlocfilehash: 6fb1f0db839910e91d7a5215f8e21f5b33ff2165
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432196"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1a967-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1a967-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1a967-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="1a967-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="1a967-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="1a967-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="1a967-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="1a967-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="1a967-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="1a967-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="1a967-108">Excel</span><span class="sxs-lookup"><span data-stu-id="1a967-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1a967-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1a967-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1a967-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1a967-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="1a967-113">Office Online</span></span></td>
    <td> <span data-ttu-id="1a967-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-114">- TaskPane</span></span><br><span data-ttu-id="1a967-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-115">
        - Content</span></span><br><span data-ttu-id="1a967-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-116">
        - Custom Functions</span></span><br><span data-ttu-id="1a967-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="1a967-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1a967-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1a967-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1a967-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1a967-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1a967-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1a967-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1a967-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1a967-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1a967-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1a967-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1a967-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1a967-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-128">
        - BindingEvents</span></span><br><span data-ttu-id="1a967-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-129">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-130">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-131">
        - File</span></span><br><span data-ttu-id="1a967-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-132">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-134">
        - Selection</span></span><br><span data-ttu-id="1a967-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-135">
        - Settings</span></span><br><span data-ttu-id="1a967-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-136">
        - TableBindings</span></span><br><span data-ttu-id="1a967-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-137">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-138">
        - TextBindings</span></span><br><span data-ttu-id="1a967-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-140">Office on Windows</span></span><br><span data-ttu-id="1a967-141">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-142">- TaskPane</span></span><br><span data-ttu-id="1a967-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-143">
        - Content</span></span><br><span data-ttu-id="1a967-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-144">
        - Custom Functions</span></span><br><span data-ttu-id="1a967-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="1a967-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1a967-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1a967-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1a967-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1a967-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1a967-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1a967-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1a967-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1a967-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1a967-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1a967-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1a967-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1a967-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-156">
        - BindingEvents</span></span><br><span data-ttu-id="1a967-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-157">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-158">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-159">
        - File</span></span><br><span data-ttu-id="1a967-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-160">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-162">
        - Selection</span></span><br><span data-ttu-id="1a967-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-163">
        - Settings</span></span><br><span data-ttu-id="1a967-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-164">
        - TableBindings</span></span><br><span data-ttu-id="1a967-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-165">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-166">
        - TextBindings</span></span><br><span data-ttu-id="1a967-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-168">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-168">Office 2019 on Windows</span></span><br><span data-ttu-id="1a967-169">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1a967-170">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-170">- TaskPane</span></span><br><span data-ttu-id="1a967-171">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-171">
        - Content</span></span><br><span data-ttu-id="1a967-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1a967-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1a967-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1a967-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1a967-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1a967-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1a967-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1a967-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1a967-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1a967-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1a967-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-182">- BindingEvents</span></span><br><span data-ttu-id="1a967-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-183">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-184">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-185">
        - File</span></span><br><span data-ttu-id="1a967-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-186">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-187">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-189">
        - Selection</span></span><br><span data-ttu-id="1a967-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-190">
        - Settings</span></span><br><span data-ttu-id="1a967-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-191">
        - TableBindings</span></span><br><span data-ttu-id="1a967-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-192">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-193">
        - TextBindings</span></span><br><span data-ttu-id="1a967-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-195">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-195">Office 2016 on Windows</span></span><br><span data-ttu-id="1a967-196">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1a967-197">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-197">- TaskPane</span></span><br><span data-ttu-id="1a967-198">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-198">
        - Content</span></span></td>
    <td><span data-ttu-id="1a967-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="1a967-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-201">- BindingEvents</span></span><br><span data-ttu-id="1a967-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-202">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-203">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-204">
        - File</span></span><br><span data-ttu-id="1a967-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-205">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-206">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-208">
        - Selection</span></span><br><span data-ttu-id="1a967-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-209">
        - Settings</span></span><br><span data-ttu-id="1a967-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-210">
        - TableBindings</span></span><br><span data-ttu-id="1a967-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-211">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-212">
        - TextBindings</span></span><br><span data-ttu-id="1a967-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-214">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-214">Office 2013 on Windows</span></span><br><span data-ttu-id="1a967-215">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1a967-216">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-216">
        - TaskPane</span></span><br><span data-ttu-id="1a967-217">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1a967-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1a967-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="1a967-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-219">
        - BindingEvents</span></span><br><span data-ttu-id="1a967-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-220">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-221">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-222">
        - File</span></span><br><span data-ttu-id="1a967-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-223">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-224">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-226">
        - Selection</span></span><br><span data-ttu-id="1a967-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-227">
        - Settings</span></span><br><span data-ttu-id="1a967-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-228">
        - TableBindings</span></span><br><span data-ttu-id="1a967-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-229">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-230">
        - TextBindings</span></span><br><span data-ttu-id="1a967-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-232">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="1a967-232">Office for iPad</span></span><br><span data-ttu-id="1a967-233">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1a967-234">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-234">- TaskPane</span></span><br><span data-ttu-id="1a967-235">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-235">
        - Content</span></span><br><span data-ttu-id="1a967-236">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1a967-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1a967-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1a967-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1a967-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1a967-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1a967-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1a967-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1a967-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1a967-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1a967-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1a967-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1a967-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-247">- BindingEvents</span></span><br><span data-ttu-id="1a967-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-248">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-249">
        - File</span></span><br><span data-ttu-id="1a967-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-250">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-251">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-253">
        - Selection</span></span><br><span data-ttu-id="1a967-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-254">
        - Settings</span></span><br><span data-ttu-id="1a967-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-255">
        - TableBindings</span></span><br><span data-ttu-id="1a967-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-256">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-257">
        - TextBindings</span></span><br><span data-ttu-id="1a967-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-259">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-259">Office for Mac</span></span><br><span data-ttu-id="1a967-260">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1a967-261">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-261">- TaskPane</span></span><br><span data-ttu-id="1a967-262">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-262">
        - Content</span></span><br><span data-ttu-id="1a967-263">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-263">
        - Custom Functions</span></span><br><span data-ttu-id="1a967-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1a967-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1a967-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1a967-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1a967-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1a967-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1a967-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1a967-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1a967-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1a967-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1a967-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1a967-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1a967-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-275">- BindingEvents</span></span><br><span data-ttu-id="1a967-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-276">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-277">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-278">
        - File</span></span><br><span data-ttu-id="1a967-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-279">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-280">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-282">
        - PdfFile</span></span><br><span data-ttu-id="1a967-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-283">
        - Selection</span></span><br><span data-ttu-id="1a967-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-284">
        - Settings</span></span><br><span data-ttu-id="1a967-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-285">
        - TableBindings</span></span><br><span data-ttu-id="1a967-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-286">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-287">
        - TextBindings</span></span><br><span data-ttu-id="1a967-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-289">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-289">Office 2019 for Mac</span></span><br><span data-ttu-id="1a967-290">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1a967-291">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-291">- TaskPane</span></span><br><span data-ttu-id="1a967-292">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-292">
        - Content</span></span><br><span data-ttu-id="1a967-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1a967-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1a967-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1a967-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1a967-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1a967-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1a967-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1a967-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1a967-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1a967-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1a967-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-303">- BindingEvents</span></span><br><span data-ttu-id="1a967-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-304">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-305">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-306">
        - File</span></span><br><span data-ttu-id="1a967-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-307">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-308">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-310">
        - PdfFile</span></span><br><span data-ttu-id="1a967-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-311">
        - Selection</span></span><br><span data-ttu-id="1a967-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-312">
        - Settings</span></span><br><span data-ttu-id="1a967-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-313">
        - TableBindings</span></span><br><span data-ttu-id="1a967-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-314">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-315">
        - TextBindings</span></span><br><span data-ttu-id="1a967-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-317">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-317">Office 2016 for Mac</span></span><br><span data-ttu-id="1a967-318">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1a967-319">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-319">- TaskPane</span></span><br><span data-ttu-id="1a967-320">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-320">
        - Content</span></span></td>
    <td><span data-ttu-id="1a967-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1a967-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="1a967-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-323">- BindingEvents</span></span><br><span data-ttu-id="1a967-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-324">
        - CompressedFile</span></span><br><span data-ttu-id="1a967-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-325">
        - DocumentEvents</span></span><br><span data-ttu-id="1a967-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="1a967-326">
        - File</span></span><br><span data-ttu-id="1a967-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-327">
        - ImageCoercion</span></span><br><span data-ttu-id="1a967-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-328">
        - MatrixBindings</span></span><br><span data-ttu-id="1a967-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="1a967-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-330">
        - PdfFile</span></span><br><span data-ttu-id="1a967-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-331">
        - Selection</span></span><br><span data-ttu-id="1a967-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-332">
        - Settings</span></span><br><span data-ttu-id="1a967-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-333">
        - TableBindings</span></span><br><span data-ttu-id="1a967-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-334">
        - TableCoercion</span></span><br><span data-ttu-id="1a967-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-335">
        - TextBindings</span></span><br><span data-ttu-id="1a967-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1a967-337">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="1a967-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="1a967-338">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1a967-339">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1a967-340">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1a967-341">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1a967-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="1a967-343">Office Online</span></span></td>
    <td><span data-ttu-id="1a967-344">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1a967-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-346">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-346">Office on Windows</span></span><br><span data-ttu-id="1a967-347">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1a967-348">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1a967-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-350">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="1a967-350">Office for iPad</span></span><br><span data-ttu-id="1a967-351">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1a967-352">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1a967-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-354">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-354">Office for Mac</span></span><br><span data-ttu-id="1a967-355">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1a967-356">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="1a967-356">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1a967-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="1a967-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="1a967-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1a967-359">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-359">Platform</span></span></th>
    <th><span data-ttu-id="1a967-360">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-360">Extension points</span></span></th>
    <th><span data-ttu-id="1a967-361">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="1a967-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="1a967-363">Office Online</span></span></td>
    <td> <span data-ttu-id="1a967-364">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-364">- Mail Read</span></span><br><span data-ttu-id="1a967-365">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-365">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1a967-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1a967-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1a967-374">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-375">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-375">Office on Windows</span></span><br><span data-ttu-id="1a967-376">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-377">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-377">- Mail Read</span></span><br><span data-ttu-id="1a967-378">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-378">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1a967-380">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="1a967-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1a967-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1a967-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1a967-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1a967-388">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-389">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-389">Office 2019 on Windows</span></span><br><span data-ttu-id="1a967-390">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-391">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-391">- Mail Read</span></span><br><span data-ttu-id="1a967-392">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-392">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1a967-394">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="1a967-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1a967-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1a967-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1a967-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1a967-402">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-403">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-403">Office 2016 on Windows</span></span><br><span data-ttu-id="1a967-404">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-405">- Mail Read</span></span><br><span data-ttu-id="1a967-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-406">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1a967-408">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="1a967-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1a967-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1a967-413">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-414">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-414">Office 2013 on Windows</span></span><br><span data-ttu-id="1a967-415">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-416">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-416">- Mail Read</span></span><br><span data-ttu-id="1a967-417">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="1a967-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="1a967-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1a967-422">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-423">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="1a967-423">Office for iOS</span></span><br><span data-ttu-id="1a967-424">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-425">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-425">- Mail Read</span></span><br><span data-ttu-id="1a967-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1a967-432">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-433">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-433">Office for Mac</span></span><br><span data-ttu-id="1a967-434">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-435">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-435">- Mail Read</span></span><br><span data-ttu-id="1a967-436">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-436">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1a967-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1a967-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1a967-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1a967-445">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-446">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-446">Office 2019 for Mac</span></span><br><span data-ttu-id="1a967-447">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-448">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-448">- Mail Read</span></span><br><span data-ttu-id="1a967-449">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-449">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1a967-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1a967-457">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-458">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-458">Office 2016 for Mac</span></span><br><span data-ttu-id="1a967-459">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-460">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-460">- Mail Read</span></span><br><span data-ttu-id="1a967-461">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="1a967-461">
      - Mail Compose</span></span><br><span data-ttu-id="1a967-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1a967-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1a967-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1a967-469">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-470">Office для Android</span><span class="sxs-lookup"><span data-stu-id="1a967-470">Office for Android</span></span><br><span data-ttu-id="1a967-471">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-471">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-472">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="1a967-472">- Mail Read</span></span><br><span data-ttu-id="1a967-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1a967-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1a967-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1a967-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1a967-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1a967-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1a967-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1a967-479">Недоступно</span><span class="sxs-lookup"><span data-stu-id="1a967-479">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="1a967-480">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="1a967-480">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="1a967-481">Word</span><span class="sxs-lookup"><span data-stu-id="1a967-481">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1a967-482">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-482">Platform</span></span></th>
    <th><span data-ttu-id="1a967-483">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-483">Extension points</span></span></th>
    <th><span data-ttu-id="1a967-484">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-484">API requirement sets</span></span></th>
    <th><span data-ttu-id="1a967-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-486">Office Online</span><span class="sxs-lookup"><span data-stu-id="1a967-486">Office Online</span></span></td>
    <td> <span data-ttu-id="1a967-487">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-487">- TaskPane</span></span><br><span data-ttu-id="1a967-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1a967-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1a967-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-493">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-493">- BindingEvents</span></span><br><span data-ttu-id="1a967-494">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-494">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-495">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-495">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-496">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-496">
         - File</span></span><br><span data-ttu-id="1a967-497">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-497">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-498">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-498">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-499">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-499">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-500">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-500">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-501">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-501">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-502">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-502">
         - PdfFile</span></span><br><span data-ttu-id="1a967-503">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-503">
         - Selection</span></span><br><span data-ttu-id="1a967-504">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-504">
         - Settings</span></span><br><span data-ttu-id="1a967-505">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-505">
         - TableBindings</span></span><br><span data-ttu-id="1a967-506">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-506">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-507">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-507">
         - TextBindings</span></span><br><span data-ttu-id="1a967-508">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-508">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-509">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-509">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-510">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-510">Office on Windows</span></span><br><span data-ttu-id="1a967-511">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-511">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-512">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-512">- TaskPane</span></span><br><span data-ttu-id="1a967-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1a967-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1a967-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-518">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-518">- BindingEvents</span></span><br><span data-ttu-id="1a967-519">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-519">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-520">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-520">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-521">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-521">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-522">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-522">
         - File</span></span><br><span data-ttu-id="1a967-523">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-523">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-524">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-524">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-525">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-525">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-526">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-526">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-527">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-527">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-528">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-528">
         - PdfFile</span></span><br><span data-ttu-id="1a967-529">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-529">
         - Selection</span></span><br><span data-ttu-id="1a967-530">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-530">
         - Settings</span></span><br><span data-ttu-id="1a967-531">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-531">
         - TableBindings</span></span><br><span data-ttu-id="1a967-532">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-532">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-533">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-533">
         - TextBindings</span></span><br><span data-ttu-id="1a967-534">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-534">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-535">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-535">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-536">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-536">Office 2019 on Windows</span></span><br><span data-ttu-id="1a967-537">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-537">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-538">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-538">- TaskPane</span></span><br><span data-ttu-id="1a967-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1a967-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1a967-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-544">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-544">- BindingEvents</span></span><br><span data-ttu-id="1a967-545">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-545">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-546">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-546">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-547">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-547">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-548">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-548">
         - File</span></span><br><span data-ttu-id="1a967-549">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-549">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-550">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-550">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-551">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-551">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-552">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-552">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-553">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-553">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-554">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-554">
         - PdfFile</span></span><br><span data-ttu-id="1a967-555">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-555">
         - Selection</span></span><br><span data-ttu-id="1a967-556">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-556">
         - Settings</span></span><br><span data-ttu-id="1a967-557">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-557">
         - TableBindings</span></span><br><span data-ttu-id="1a967-558">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-558">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-559">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-559">
         - TextBindings</span></span><br><span data-ttu-id="1a967-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-560">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-561">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-561">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-562">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-562">Office 2016 on Windows</span></span><br><span data-ttu-id="1a967-563">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-563">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-564">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-564">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="1a967-567">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-567">- BindingEvents</span></span><br><span data-ttu-id="1a967-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-568">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-569">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-569">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-570">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-570">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-571">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-571">
         - File</span></span><br><span data-ttu-id="1a967-572">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-572">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-573">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-574">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-577">
         - PdfFile</span></span><br><span data-ttu-id="1a967-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-578">
         - Selection</span></span><br><span data-ttu-id="1a967-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-579">
         - Settings</span></span><br><span data-ttu-id="1a967-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-580">
         - TableBindings</span></span><br><span data-ttu-id="1a967-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-581">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-582">
         - TextBindings</span></span><br><span data-ttu-id="1a967-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-583">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-585">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-585">Office 2013 on Windows</span></span><br><span data-ttu-id="1a967-586">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-587">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1a967-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1a967-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-589">- BindingEvents</span></span><br><span data-ttu-id="1a967-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-590">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-592">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-593">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-593">
         - File</span></span><br><span data-ttu-id="1a967-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-595">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-596">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-599">
         - PdfFile</span></span><br><span data-ttu-id="1a967-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-600">
         - Selection</span></span><br><span data-ttu-id="1a967-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-601">
         - Settings</span></span><br><span data-ttu-id="1a967-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-602">
         - TableBindings</span></span><br><span data-ttu-id="1a967-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-603">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-604">
         - TextBindings</span></span><br><span data-ttu-id="1a967-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-605">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-606">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-607">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="1a967-607">Office for iPad</span></span><br><span data-ttu-id="1a967-608">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-608">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-609">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1a967-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1a967-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1a967-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1a967-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-614">- BindingEvents</span></span><br><span data-ttu-id="1a967-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-615">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-617">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-618">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-618">
         - File</span></span><br><span data-ttu-id="1a967-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-620">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-621">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-624">
         - PdfFile</span></span><br><span data-ttu-id="1a967-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-625">
         - Selection</span></span><br><span data-ttu-id="1a967-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-626">
         - Settings</span></span><br><span data-ttu-id="1a967-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-627">
         - TableBindings</span></span><br><span data-ttu-id="1a967-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-628">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-629">
         - TextBindings</span></span><br><span data-ttu-id="1a967-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-630">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-632">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-632">Office for Mac</span></span><br><span data-ttu-id="1a967-633">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-633">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-634">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-634">- TaskPane</span></span><br><span data-ttu-id="1a967-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1a967-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1a967-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1a967-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1a967-640">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-640">- BindingEvents</span></span><br><span data-ttu-id="1a967-641">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-641">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-642">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-642">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-643">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-643">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-644">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-644">
         - File</span></span><br><span data-ttu-id="1a967-645">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-645">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-646">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-646">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-647">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-647">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-648">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-648">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-649">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-649">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-650">
         - PdfFile</span></span><br><span data-ttu-id="1a967-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-651">
         - Selection</span></span><br><span data-ttu-id="1a967-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-652">
         - Settings</span></span><br><span data-ttu-id="1a967-653">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-653">
         - TableBindings</span></span><br><span data-ttu-id="1a967-654">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-654">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-655">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-655">
         - TextBindings</span></span><br><span data-ttu-id="1a967-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-656">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-657">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-657">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-658">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-658">Office 2019 for Mac</span></span><br><span data-ttu-id="1a967-659">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-659">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-660">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-660">- TaskPane</span></span><br><span data-ttu-id="1a967-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1a967-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1a967-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1a967-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1a967-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1a967-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1a967-666">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-666">- BindingEvents</span></span><br><span data-ttu-id="1a967-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-667">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-668">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-668">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-669">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-669">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-670">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-670">
         - File</span></span><br><span data-ttu-id="1a967-671">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-671">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-672">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-672">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-673">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-673">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-674">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-674">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-675">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-675">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-676">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-676">
         - PdfFile</span></span><br><span data-ttu-id="1a967-677">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-677">
         - Selection</span></span><br><span data-ttu-id="1a967-678">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-678">
         - Settings</span></span><br><span data-ttu-id="1a967-679">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-679">
         - TableBindings</span></span><br><span data-ttu-id="1a967-680">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-680">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-681">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-681">
         - TextBindings</span></span><br><span data-ttu-id="1a967-682">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-682">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-683">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-683">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-684">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-684">Office 2016 for Mac</span></span><br><span data-ttu-id="1a967-685">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-685">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-686">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1a967-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1a967-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="1a967-689">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-689">- BindingEvents</span></span><br><span data-ttu-id="1a967-690">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-690">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-691">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1a967-691">
         - CustomXmlParts</span></span><br><span data-ttu-id="1a967-692">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-692">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-693">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="1a967-693">
         - File</span></span><br><span data-ttu-id="1a967-694">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-694">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-695">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-695">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-696">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-696">
         - MatrixBindings</span></span><br><span data-ttu-id="1a967-697">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-697">
         - MatrixCoercion</span></span><br><span data-ttu-id="1a967-698">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-698">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1a967-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-699">
         - PdfFile</span></span><br><span data-ttu-id="1a967-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-700">
         - Selection</span></span><br><span data-ttu-id="1a967-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1a967-701">
         - Settings</span></span><br><span data-ttu-id="1a967-702">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-702">
         - TableBindings</span></span><br><span data-ttu-id="1a967-703">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-703">
         - TableCoercion</span></span><br><span data-ttu-id="1a967-704">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1a967-704">
         - TextBindings</span></span><br><span data-ttu-id="1a967-705">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-705">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-706">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1a967-706">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="1a967-707">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="1a967-707">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1a967-708">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1a967-708">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1a967-709">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-709">Platform</span></span></th>
    <th><span data-ttu-id="1a967-710">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-710">Extension points</span></span></th>
    <th><span data-ttu-id="1a967-711">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-711">API requirement sets</span></span></th>
    <th><span data-ttu-id="1a967-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-713">Office Online</span><span class="sxs-lookup"><span data-stu-id="1a967-713">Office Online</span></span></td>
    <td> <span data-ttu-id="1a967-714">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-714">- Content</span></span><br><span data-ttu-id="1a967-715">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-715">
         - TaskPane</span></span><br><span data-ttu-id="1a967-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-718">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-718">- ActiveView</span></span><br><span data-ttu-id="1a967-719">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-719">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-720">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-720">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-721">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-721">
         - File</span></span><br><span data-ttu-id="1a967-722">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-722">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-723">
         - PdfFile</span></span><br><span data-ttu-id="1a967-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-724">
         - Selection</span></span><br><span data-ttu-id="1a967-725">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-725">
         - Settings</span></span><br><span data-ttu-id="1a967-726">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-726">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-727">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-727">Office on Windows</span></span><br><span data-ttu-id="1a967-728">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-728">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-729">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-729">- Content</span></span><br><span data-ttu-id="1a967-730">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-730">
         - TaskPane</span></span><br><span data-ttu-id="1a967-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-733">- ActiveView</span></span><br><span data-ttu-id="1a967-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-734">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-735">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-736">
         - File</span></span><br><span data-ttu-id="1a967-737">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-737">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-738">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-738">
         - PdfFile</span></span><br><span data-ttu-id="1a967-739">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-739">
         - Selection</span></span><br><span data-ttu-id="1a967-740">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-740">
         - Settings</span></span><br><span data-ttu-id="1a967-741">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-741">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-742">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-742">Office 2019 on Windows</span></span><br><span data-ttu-id="1a967-743">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-743">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-744">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-744">- Content</span></span><br><span data-ttu-id="1a967-745">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-745">
         - TaskPane</span></span><br><span data-ttu-id="1a967-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-748">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-748">- ActiveView</span></span><br><span data-ttu-id="1a967-749">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-749">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-750">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-750">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-751">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-751">
         - File</span></span><br><span data-ttu-id="1a967-752">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-752">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-753">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-753">
         - PdfFile</span></span><br><span data-ttu-id="1a967-754">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-754">
         - Selection</span></span><br><span data-ttu-id="1a967-755">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-755">
         - Settings</span></span><br><span data-ttu-id="1a967-756">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-756">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-757">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-757">Office 2016 on Windows</span></span><br><span data-ttu-id="1a967-758">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-758">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-759">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-759">- Content</span></span><br><span data-ttu-id="1a967-760">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-760">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1a967-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1a967-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-762">- ActiveView</span></span><br><span data-ttu-id="1a967-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-763">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-764">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-765">
         - File</span></span><br><span data-ttu-id="1a967-766">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-766">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-767">
         - PdfFile</span></span><br><span data-ttu-id="1a967-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-768">
         - Selection</span></span><br><span data-ttu-id="1a967-769">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-769">
         - Settings</span></span><br><span data-ttu-id="1a967-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-771">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-771">Office 2013 on Windows</span></span><br><span data-ttu-id="1a967-772">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-772">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-773">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-773">- Content</span></span><br><span data-ttu-id="1a967-774">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-774">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="1a967-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1a967-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1a967-776">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-776">- ActiveView</span></span><br><span data-ttu-id="1a967-777">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-777">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-778">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-778">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-779">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-779">
         - File</span></span><br><span data-ttu-id="1a967-780">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-780">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-781">
         - PdfFile</span></span><br><span data-ttu-id="1a967-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-782">
         - Selection</span></span><br><span data-ttu-id="1a967-783">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-783">
         - Settings</span></span><br><span data-ttu-id="1a967-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-785">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="1a967-785">Office for iPad</span></span><br><span data-ttu-id="1a967-786">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-786">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-787">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-787">- Content</span></span><br><span data-ttu-id="1a967-788">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-790">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-790">- ActiveView</span></span><br><span data-ttu-id="1a967-791">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-791">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-792">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-792">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-793">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-793">
         - File</span></span><br><span data-ttu-id="1a967-794">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-794">
         - PdfFile</span></span><br><span data-ttu-id="1a967-795">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-795">
         - Selection</span></span><br><span data-ttu-id="1a967-796">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-796">
         - Settings</span></span><br><span data-ttu-id="1a967-797">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-797">
         - TextCoercion</span></span><br><span data-ttu-id="1a967-798">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-798">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-799">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-799">Office for Mac</span></span><br><span data-ttu-id="1a967-800">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="1a967-800">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="1a967-801">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-801">- Content</span></span><br><span data-ttu-id="1a967-802">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-802">
         - TaskPane</span></span><br><span data-ttu-id="1a967-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-805">- ActiveView</span></span><br><span data-ttu-id="1a967-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-806">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-807">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-808">
         - File</span></span><br><span data-ttu-id="1a967-809">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-809">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-810">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-810">
         - PdfFile</span></span><br><span data-ttu-id="1a967-811">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-811">
         - Selection</span></span><br><span data-ttu-id="1a967-812">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-812">
         - Settings</span></span><br><span data-ttu-id="1a967-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-814">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-814">Office 2019 for Mac</span></span><br><span data-ttu-id="1a967-815">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-815">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-816">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-816">- Content</span></span><br><span data-ttu-id="1a967-817">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-817">
         - TaskPane</span></span><br><span data-ttu-id="1a967-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-820">- ActiveView</span></span><br><span data-ttu-id="1a967-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-821">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-822">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-823">
         - File</span></span><br><span data-ttu-id="1a967-824">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-824">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-825">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-825">
         - PdfFile</span></span><br><span data-ttu-id="1a967-826">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-826">
         - Selection</span></span><br><span data-ttu-id="1a967-827">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-827">
         - Settings</span></span><br><span data-ttu-id="1a967-828">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-828">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-829">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-829">Office 2016 for Mac</span></span><br><span data-ttu-id="1a967-830">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-830">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-831">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-831">- Content</span></span><br><span data-ttu-id="1a967-832">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-832">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1a967-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1a967-834">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1a967-834">- ActiveView</span></span><br><span data-ttu-id="1a967-835">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1a967-835">
         - CompressedFile</span></span><br><span data-ttu-id="1a967-836">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-836">
         - DocumentEvents</span></span><br><span data-ttu-id="1a967-837">
         - File</span><span class="sxs-lookup"><span data-stu-id="1a967-837">
         - File</span></span><br><span data-ttu-id="1a967-838">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-838">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-839">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1a967-839">
         - PdfFile</span></span><br><span data-ttu-id="1a967-840">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-840">
         - Selection</span></span><br><span data-ttu-id="1a967-841">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-841">
         - Settings</span></span><br><span data-ttu-id="1a967-842">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-842">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1a967-843">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="1a967-843">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="1a967-844">OneNote</span><span class="sxs-lookup"><span data-stu-id="1a967-844">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1a967-845">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-845">Platform</span></span></th>
    <th><span data-ttu-id="1a967-846">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-846">Extension points</span></span></th>
    <th><span data-ttu-id="1a967-847">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-847">API requirement sets</span></span></th>
    <th><span data-ttu-id="1a967-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-849">Office Online</span><span class="sxs-lookup"><span data-stu-id="1a967-849">Office Online</span></span></td>
    <td> <span data-ttu-id="1a967-850">- Контент</span><span class="sxs-lookup"><span data-stu-id="1a967-850">- Content</span></span><br><span data-ttu-id="1a967-851">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-851">
         - TaskPane</span></span><br><span data-ttu-id="1a967-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="1a967-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1a967-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1a967-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-855">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1a967-855">- DocumentEvents</span></span><br><span data-ttu-id="1a967-856">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-856">
         - HtmlCoercion</span></span><br><span data-ttu-id="1a967-857">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-857">
         - ImageCoercion</span></span><br><span data-ttu-id="1a967-858">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="1a967-858">
         - Settings</span></span><br><span data-ttu-id="1a967-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-859">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="1a967-860">Project</span><span class="sxs-lookup"><span data-stu-id="1a967-860">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1a967-861">Платформа</span><span class="sxs-lookup"><span data-stu-id="1a967-861">Platform</span></span></th>
    <th><span data-ttu-id="1a967-862">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="1a967-862">Extension points</span></span></th>
    <th><span data-ttu-id="1a967-863">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="1a967-863">API requirement sets</span></span></th>
    <th><span data-ttu-id="1a967-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="1a967-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-865">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-865">Office 2019 on Windows</span></span><br><span data-ttu-id="1a967-866">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-866">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-867">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-867">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-869">- Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-869">- Selection</span></span><br><span data-ttu-id="1a967-870">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-870">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-871">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-871">Office 2016 on Windows</span></span><br><span data-ttu-id="1a967-872">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-872">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-873">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-873">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-875">- Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-875">- Selection</span></span><br><span data-ttu-id="1a967-876">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-876">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1a967-877">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="1a967-877">Office 2013 on Windows</span></span><br><span data-ttu-id="1a967-878">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="1a967-878">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1a967-879">- Область задач</span><span class="sxs-lookup"><span data-stu-id="1a967-879">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1a967-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1a967-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1a967-881">- Selection</span><span class="sxs-lookup"><span data-stu-id="1a967-881">- Selection</span></span><br><span data-ttu-id="1a967-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1a967-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1a967-883">См. также</span><span class="sxs-lookup"><span data-stu-id="1a967-883">See also</span></span>

- [<span data-ttu-id="1a967-884">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1a967-884">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1a967-885">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="1a967-885">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="1a967-886">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="1a967-886">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="1a967-887">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="1a967-887">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="1a967-888">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="1a967-888">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="1a967-889">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="1a967-889">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="1a967-890">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="1a967-890">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="1a967-891">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="1a967-891">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="1a967-892">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="1a967-892">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="1a967-893">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="1a967-893">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="1a967-894">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="1a967-894">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
