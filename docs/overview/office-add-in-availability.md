---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448149"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="52f90-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="52f90-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="52f90-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="52f90-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="52f90-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="52f90-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="52f90-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="52f90-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="52f90-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="52f90-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="52f90-108">Excel</span><span class="sxs-lookup"><span data-stu-id="52f90-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="52f90-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="52f90-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="52f90-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="52f90-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="52f90-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="52f90-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="52f90-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="52f90-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="52f90-113">Office Online</span></span></td>
    <td> <span data-ttu-id="52f90-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-114">- TaskPane</span></span><br><span data-ttu-id="52f90-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-115">
        - Content</span></span><br><span data-ttu-id="52f90-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="52f90-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="52f90-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52f90-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52f90-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52f90-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52f90-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52f90-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="52f90-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52f90-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52f90-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52f90-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-126">
        - BindingEvents</span></span><br><span data-ttu-id="52f90-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-127">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-128">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-129">
        - File</span></span><br><span data-ttu-id="52f90-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-130">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-132">
        - Selection</span></span><br><span data-ttu-id="52f90-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-133">
        - Settings</span></span><br><span data-ttu-id="52f90-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-134">
        - TableBindings</span></span><br><span data-ttu-id="52f90-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-135">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-136">
        - TextBindings</span></span><br><span data-ttu-id="52f90-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-138">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-138">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-139">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-139">- TaskPane</span></span><br><span data-ttu-id="52f90-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-140">
        - Content</span></span><br><span data-ttu-id="52f90-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="52f90-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="52f90-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52f90-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52f90-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52f90-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52f90-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52f90-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="52f90-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52f90-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52f90-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52f90-151">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-151">
        - BindingEvents</span></span><br><span data-ttu-id="52f90-152">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-152">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-153">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-153">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-154">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-154">
        - File</span></span><br><span data-ttu-id="52f90-155">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-155">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-156">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-156">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-157">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-157">
        - Selection</span></span><br><span data-ttu-id="52f90-158">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-158">
        - Settings</span></span><br><span data-ttu-id="52f90-159">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-159">
        - TableBindings</span></span><br><span data-ttu-id="52f90-160">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-160">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-161">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-161">
        - TextBindings</span></span><br><span data-ttu-id="52f90-162">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-162">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-163">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-163">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="52f90-164">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-164">- TaskPane</span></span><br><span data-ttu-id="52f90-165">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-165">
        - Content</span></span><br><span data-ttu-id="52f90-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52f90-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52f90-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52f90-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52f90-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52f90-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52f90-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="52f90-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52f90-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52f90-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52f90-176">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-176">- BindingEvents</span></span><br><span data-ttu-id="52f90-177">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-177">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-178">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-178">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-179">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-179">
        - File</span></span><br><span data-ttu-id="52f90-180">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-180">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-181">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-181">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-182">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-182">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-183">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-183">
        - Selection</span></span><br><span data-ttu-id="52f90-184">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-184">
        - Settings</span></span><br><span data-ttu-id="52f90-185">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-185">
        - TableBindings</span></span><br><span data-ttu-id="52f90-186">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-186">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-187">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-187">
        - TextBindings</span></span><br><span data-ttu-id="52f90-188">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-188">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-189">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-189">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="52f90-190">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-190">- TaskPane</span></span><br><span data-ttu-id="52f90-191">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-191">
        - Content</span></span></td>
    <td><span data-ttu-id="52f90-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="52f90-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-194">- BindingEvents</span></span><br><span data-ttu-id="52f90-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-195">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-196">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-197">
        - File</span></span><br><span data-ttu-id="52f90-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-198">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-199">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-201">
        - Selection</span></span><br><span data-ttu-id="52f90-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-202">
        - Settings</span></span><br><span data-ttu-id="52f90-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-203">
        - TableBindings</span></span><br><span data-ttu-id="52f90-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-204">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-205">
        - TextBindings</span></span><br><span data-ttu-id="52f90-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-207">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-207">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="52f90-208">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-208">
        - TaskPane</span></span><br><span data-ttu-id="52f90-209">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-209">
        - Content</span></span></td>
    <td>  <span data-ttu-id="52f90-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="52f90-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="52f90-211">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-211">
        - BindingEvents</span></span><br><span data-ttu-id="52f90-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-212">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-213">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-214">
        - File</span></span><br><span data-ttu-id="52f90-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-215">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-216">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-218">
        - Selection</span></span><br><span data-ttu-id="52f90-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-219">
        - Settings</span></span><br><span data-ttu-id="52f90-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-220">
        - TableBindings</span></span><br><span data-ttu-id="52f90-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-221">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-222">
        - TextBindings</span></span><br><span data-ttu-id="52f90-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-224">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="52f90-224">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="52f90-225">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-225">- TaskPane</span></span><br><span data-ttu-id="52f90-226">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-226">
        - Content</span></span></td>
    <td><span data-ttu-id="52f90-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52f90-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52f90-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52f90-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52f90-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52f90-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="52f90-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52f90-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52f90-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52f90-236">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-236">- BindingEvents</span></span><br><span data-ttu-id="52f90-237">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-237">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-238">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-238">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-239">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-239">
        - File</span></span><br><span data-ttu-id="52f90-240">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-240">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-241">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-241">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-242">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-242">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-243">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-243">
        - Selection</span></span><br><span data-ttu-id="52f90-244">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-244">
        - Settings</span></span><br><span data-ttu-id="52f90-245">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-245">
        - TableBindings</span></span><br><span data-ttu-id="52f90-246">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-246">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-247">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-247">
        - TextBindings</span></span><br><span data-ttu-id="52f90-248">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-248">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-249">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-249">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="52f90-250">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-250">- TaskPane</span></span><br><span data-ttu-id="52f90-251">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-251">
        - Content</span></span><br><span data-ttu-id="52f90-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52f90-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52f90-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52f90-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52f90-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52f90-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52f90-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="52f90-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52f90-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52f90-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52f90-262">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-262">- BindingEvents</span></span><br><span data-ttu-id="52f90-263">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-263">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-264">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-264">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-265">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-265">
        - File</span></span><br><span data-ttu-id="52f90-266">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-266">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-267">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-267">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-268">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-268">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-269">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-269">
        - PdfFile</span></span><br><span data-ttu-id="52f90-270">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-270">
        - Selection</span></span><br><span data-ttu-id="52f90-271">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-271">
        - Settings</span></span><br><span data-ttu-id="52f90-272">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-272">
        - TableBindings</span></span><br><span data-ttu-id="52f90-273">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-273">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-274">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-274">
        - TextBindings</span></span><br><span data-ttu-id="52f90-275">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-275">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-276">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-276">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="52f90-277">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-277">- TaskPane</span></span><br><span data-ttu-id="52f90-278">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-278">
        - Content</span></span><br><span data-ttu-id="52f90-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="52f90-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="52f90-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="52f90-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="52f90-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="52f90-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="52f90-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="52f90-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="52f90-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="52f90-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="52f90-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-289">- BindingEvents</span></span><br><span data-ttu-id="52f90-290">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-290">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-291">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-291">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-292">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-292">
        - File</span></span><br><span data-ttu-id="52f90-293">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-293">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-294">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-294">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-295">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-295">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-296">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-296">
        - PdfFile</span></span><br><span data-ttu-id="52f90-297">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-297">
        - Selection</span></span><br><span data-ttu-id="52f90-298">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-298">
        - Settings</span></span><br><span data-ttu-id="52f90-299">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-299">
        - TableBindings</span></span><br><span data-ttu-id="52f90-300">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-300">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-301">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-301">
        - TextBindings</span></span><br><span data-ttu-id="52f90-302">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-302">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-303">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-303">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="52f90-304">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-304">- TaskPane</span></span><br><span data-ttu-id="52f90-305">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-305">
        - Content</span></span></td>
    <td><span data-ttu-id="52f90-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="52f90-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="52f90-308">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-308">- BindingEvents</span></span><br><span data-ttu-id="52f90-309">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-309">
        - CompressedFile</span></span><br><span data-ttu-id="52f90-310">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-310">
        - DocumentEvents</span></span><br><span data-ttu-id="52f90-311">
        - File</span><span class="sxs-lookup"><span data-stu-id="52f90-311">
        - File</span></span><br><span data-ttu-id="52f90-312">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-312">
        - ImageCoercion</span></span><br><span data-ttu-id="52f90-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-313">
        - MatrixBindings</span></span><br><span data-ttu-id="52f90-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="52f90-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-315">
        - PdfFile</span></span><br><span data-ttu-id="52f90-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-316">
        - Selection</span></span><br><span data-ttu-id="52f90-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-317">
        - Settings</span></span><br><span data-ttu-id="52f90-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-318">
        - TableBindings</span></span><br><span data-ttu-id="52f90-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-319">
        - TableCoercion</span></span><br><span data-ttu-id="52f90-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-320">
        - TextBindings</span></span><br><span data-ttu-id="52f90-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-321">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="52f90-322">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="52f90-322">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="52f90-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="52f90-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52f90-324">Платформа</span><span class="sxs-lookup"><span data-stu-id="52f90-324">Platform</span></span></th>
    <th><span data-ttu-id="52f90-325">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="52f90-325">Extension points</span></span></th>
    <th><span data-ttu-id="52f90-326">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="52f90-326">API requirement sets</span></span></th>
    <th><span data-ttu-id="52f90-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="52f90-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="52f90-328">Office Online</span></span></td>
    <td> <span data-ttu-id="52f90-329">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-329">- Mail Read</span></span><br><span data-ttu-id="52f90-330">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-330">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52f90-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="52f90-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="52f90-339">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-340">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-340">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-341">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-341">- Mail Read</span></span><br><span data-ttu-id="52f90-342">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-342">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="52f90-344">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="52f90-344">
      - Modules</span></span></td>
    <td> <span data-ttu-id="52f90-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52f90-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="52f90-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="52f90-352">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-353">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-353">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-354">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-354">- Mail Read</span></span><br><span data-ttu-id="52f90-355">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-355">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="52f90-357">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="52f90-357">
      - Modules</span></span></td>
    <td> <span data-ttu-id="52f90-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52f90-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="52f90-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="52f90-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="52f90-365">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-366">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-366">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-367">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-367">- Mail Read</span></span><br><span data-ttu-id="52f90-368">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-368">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="52f90-370">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="52f90-370">
      - Modules</span></span></td>
    <td> <span data-ttu-id="52f90-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="52f90-375">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-376">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-376">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-377">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-377">- Mail Read</span></span><br><span data-ttu-id="52f90-378">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-378">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="52f90-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="52f90-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="52f90-383">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-384">Office 365 для iOS</span><span class="sxs-lookup"><span data-stu-id="52f90-384">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="52f90-385">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-385">- Mail Read</span></span><br><span data-ttu-id="52f90-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="52f90-392">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-393">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-393">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-394">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-394">- Mail Read</span></span><br><span data-ttu-id="52f90-395">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-395">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52f90-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="52f90-403">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-404">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-404">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-405">- Mail Read</span></span><br><span data-ttu-id="52f90-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-406">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52f90-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="52f90-414">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-415">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-415">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-416">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-416">- Mail Read</span></span><br><span data-ttu-id="52f90-417">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="52f90-417">
      - Mail Compose</span></span><br><span data-ttu-id="52f90-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="52f90-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="52f90-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="52f90-425">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-426">Office 365 для Android</span><span class="sxs-lookup"><span data-stu-id="52f90-426">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="52f90-427">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="52f90-427">- Mail Read</span></span><br><span data-ttu-id="52f90-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="52f90-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="52f90-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="52f90-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="52f90-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="52f90-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="52f90-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="52f90-434">Недоступно</span><span class="sxs-lookup"><span data-stu-id="52f90-434">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="52f90-435">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="52f90-435">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="52f90-436">Word</span><span class="sxs-lookup"><span data-stu-id="52f90-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52f90-437">Платформа</span><span class="sxs-lookup"><span data-stu-id="52f90-437">Platform</span></span></th>
    <th><span data-ttu-id="52f90-438">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="52f90-438">Extension points</span></span></th>
    <th><span data-ttu-id="52f90-439">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="52f90-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="52f90-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="52f90-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="52f90-441">Office Online</span></span></td>
    <td> <span data-ttu-id="52f90-442">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-442">- TaskPane</span></span><br><span data-ttu-id="52f90-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52f90-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52f90-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-448">- BindingEvents</span></span><br><span data-ttu-id="52f90-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-450">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-451">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-451">
         - File</span></span><br><span data-ttu-id="52f90-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-453">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-454">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-457">
         - PdfFile</span></span><br><span data-ttu-id="52f90-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-458">
         - Selection</span></span><br><span data-ttu-id="52f90-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-459">
         - Settings</span></span><br><span data-ttu-id="52f90-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-460">
         - TableBindings</span></span><br><span data-ttu-id="52f90-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-461">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-462">
         - TextBindings</span></span><br><span data-ttu-id="52f90-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-463">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-465">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-466">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-466">- TaskPane</span></span><br><span data-ttu-id="52f90-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52f90-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52f90-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-472">- BindingEvents</span></span><br><span data-ttu-id="52f90-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-473">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-475">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-476">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-476">
         - File</span></span><br><span data-ttu-id="52f90-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-478">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-479">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-482">
         - PdfFile</span></span><br><span data-ttu-id="52f90-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-483">
         - Selection</span></span><br><span data-ttu-id="52f90-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-484">
         - Settings</span></span><br><span data-ttu-id="52f90-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-485">
         - TableBindings</span></span><br><span data-ttu-id="52f90-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-486">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-487">
         - TextBindings</span></span><br><span data-ttu-id="52f90-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-488">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-490">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-491">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-491">- TaskPane</span></span><br><span data-ttu-id="52f90-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52f90-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52f90-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-497">- BindingEvents</span></span><br><span data-ttu-id="52f90-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-498">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-500">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-501">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-501">
         - File</span></span><br><span data-ttu-id="52f90-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-503">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-504">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-507">
         - PdfFile</span></span><br><span data-ttu-id="52f90-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-508">
         - Selection</span></span><br><span data-ttu-id="52f90-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-509">
         - Settings</span></span><br><span data-ttu-id="52f90-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-510">
         - TableBindings</span></span><br><span data-ttu-id="52f90-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-511">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-512">
         - TextBindings</span></span><br><span data-ttu-id="52f90-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-513">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-515">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-516">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="52f90-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-519">- BindingEvents</span></span><br><span data-ttu-id="52f90-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-520">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-522">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-523">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-523">
         - File</span></span><br><span data-ttu-id="52f90-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-525">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-526">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-529">
         - PdfFile</span></span><br><span data-ttu-id="52f90-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-530">
         - Selection</span></span><br><span data-ttu-id="52f90-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-531">
         - Settings</span></span><br><span data-ttu-id="52f90-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-532">
         - TableBindings</span></span><br><span data-ttu-id="52f90-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-533">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-534">
         - TextBindings</span></span><br><span data-ttu-id="52f90-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-535">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-537">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-538">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="52f90-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="52f90-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-540">- BindingEvents</span></span><br><span data-ttu-id="52f90-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-541">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-543">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-544">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-544">
         - File</span></span><br><span data-ttu-id="52f90-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-546">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-547">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-550">
         - PdfFile</span></span><br><span data-ttu-id="52f90-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-551">
         - Selection</span></span><br><span data-ttu-id="52f90-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-552">
         - Settings</span></span><br><span data-ttu-id="52f90-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-553">
         - TableBindings</span></span><br><span data-ttu-id="52f90-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-554">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-555">
         - TextBindings</span></span><br><span data-ttu-id="52f90-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-556">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-558">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="52f90-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="52f90-559">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52f90-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52f90-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52f90-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52f90-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-564">- BindingEvents</span></span><br><span data-ttu-id="52f90-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-565">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-567">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-568">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-568">
         - File</span></span><br><span data-ttu-id="52f90-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-570">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-571">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-574">
         - PdfFile</span></span><br><span data-ttu-id="52f90-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-575">
         - Selection</span></span><br><span data-ttu-id="52f90-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-576">
         - Settings</span></span><br><span data-ttu-id="52f90-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-577">
         - TableBindings</span></span><br><span data-ttu-id="52f90-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-578">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-579">
         - TextBindings</span></span><br><span data-ttu-id="52f90-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-580">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-582">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-583">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-583">- TaskPane</span></span><br><span data-ttu-id="52f90-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52f90-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52f90-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52f90-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52f90-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-589">- BindingEvents</span></span><br><span data-ttu-id="52f90-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-590">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-592">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-593">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-593">
         - File</span></span><br><span data-ttu-id="52f90-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-595">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-596">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-599">
         - PdfFile</span></span><br><span data-ttu-id="52f90-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-600">
         - Selection</span></span><br><span data-ttu-id="52f90-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-601">
         - Settings</span></span><br><span data-ttu-id="52f90-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-602">
         - TableBindings</span></span><br><span data-ttu-id="52f90-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-603">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-604">
         - TextBindings</span></span><br><span data-ttu-id="52f90-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-605">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-607">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-608">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-608">- TaskPane</span></span><br><span data-ttu-id="52f90-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="52f90-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="52f90-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="52f90-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="52f90-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="52f90-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="52f90-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-614">- BindingEvents</span></span><br><span data-ttu-id="52f90-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-615">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-617">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-618">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-618">
         - File</span></span><br><span data-ttu-id="52f90-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-620">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-621">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-624">
         - PdfFile</span></span><br><span data-ttu-id="52f90-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-625">
         - Selection</span></span><br><span data-ttu-id="52f90-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-626">
         - Settings</span></span><br><span data-ttu-id="52f90-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-627">
         - TableBindings</span></span><br><span data-ttu-id="52f90-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-628">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-629">
         - TextBindings</span></span><br><span data-ttu-id="52f90-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-630">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-632">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-633">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="52f90-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="52f90-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="52f90-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-636">- BindingEvents</span></span><br><span data-ttu-id="52f90-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-637">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="52f90-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="52f90-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-639">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-640">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="52f90-640">
         - File</span></span><br><span data-ttu-id="52f90-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-642">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-643">
         - MatrixBindings</span></span><br><span data-ttu-id="52f90-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="52f90-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="52f90-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-646">
         - PdfFile</span></span><br><span data-ttu-id="52f90-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-647">
         - Selection</span></span><br><span data-ttu-id="52f90-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="52f90-648">
         - Settings</span></span><br><span data-ttu-id="52f90-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-649">
         - TableBindings</span></span><br><span data-ttu-id="52f90-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-650">
         - TableCoercion</span></span><br><span data-ttu-id="52f90-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="52f90-651">
         - TextBindings</span></span><br><span data-ttu-id="52f90-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-652">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="52f90-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="52f90-654">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="52f90-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="52f90-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="52f90-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52f90-656">Платформа</span><span class="sxs-lookup"><span data-stu-id="52f90-656">Platform</span></span></th>
    <th><span data-ttu-id="52f90-657">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="52f90-657">Extension points</span></span></th>
    <th><span data-ttu-id="52f90-658">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="52f90-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="52f90-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="52f90-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="52f90-660">Office Online</span></span></td>
    <td> <span data-ttu-id="52f90-661">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-661">- Content</span></span><br><span data-ttu-id="52f90-662">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-662">
         - TaskPane</span></span><br><span data-ttu-id="52f90-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-665">- ActiveView</span></span><br><span data-ttu-id="52f90-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-666">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-667">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-668">
         - File</span></span><br><span data-ttu-id="52f90-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-669">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-670">
         - PdfFile</span></span><br><span data-ttu-id="52f90-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-671">
         - Selection</span></span><br><span data-ttu-id="52f90-672">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-672">
         - Settings</span></span><br><span data-ttu-id="52f90-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-674">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-675">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-675">- Content</span></span><br><span data-ttu-id="52f90-676">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-676">
         - TaskPane</span></span><br><span data-ttu-id="52f90-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-679">- ActiveView</span></span><br><span data-ttu-id="52f90-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-680">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-681">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-682">
         - File</span></span><br><span data-ttu-id="52f90-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-683">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-684">
         - PdfFile</span></span><br><span data-ttu-id="52f90-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-685">
         - Selection</span></span><br><span data-ttu-id="52f90-686">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-686">
         - Settings</span></span><br><span data-ttu-id="52f90-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-688">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-689">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-689">- Content</span></span><br><span data-ttu-id="52f90-690">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-690">
         - TaskPane</span></span><br><span data-ttu-id="52f90-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-693">- ActiveView</span></span><br><span data-ttu-id="52f90-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-694">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-695">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-696">
         - File</span></span><br><span data-ttu-id="52f90-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-697">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-698">
         - PdfFile</span></span><br><span data-ttu-id="52f90-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-699">
         - Selection</span></span><br><span data-ttu-id="52f90-700">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-700">
         - Settings</span></span><br><span data-ttu-id="52f90-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-702">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-703">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-703">- Content</span></span><br><span data-ttu-id="52f90-704">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="52f90-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="52f90-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-706">- ActiveView</span></span><br><span data-ttu-id="52f90-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-707">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-708">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-709">
         - File</span></span><br><span data-ttu-id="52f90-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-710">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-711">
         - PdfFile</span></span><br><span data-ttu-id="52f90-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-712">
         - Selection</span></span><br><span data-ttu-id="52f90-713">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-713">
         - Settings</span></span><br><span data-ttu-id="52f90-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-715">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-716">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-716">- Content</span></span><br><span data-ttu-id="52f90-717">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="52f90-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="52f90-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="52f90-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-719">- ActiveView</span></span><br><span data-ttu-id="52f90-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-720">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-721">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-722">
         - File</span></span><br><span data-ttu-id="52f90-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-723">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-724">
         - PdfFile</span></span><br><span data-ttu-id="52f90-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-725">
         - Selection</span></span><br><span data-ttu-id="52f90-726">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-726">
         - Settings</span></span><br><span data-ttu-id="52f90-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-728">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="52f90-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="52f90-729">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-729">- Content</span></span><br><span data-ttu-id="52f90-730">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="52f90-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-732">- ActiveView</span></span><br><span data-ttu-id="52f90-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-733">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-734">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-735">
         - File</span></span><br><span data-ttu-id="52f90-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-736">
         - PdfFile</span></span><br><span data-ttu-id="52f90-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-737">
         - Selection</span></span><br><span data-ttu-id="52f90-738">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-738">
         - Settings</span></span><br><span data-ttu-id="52f90-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-739">
         - TextCoercion</span></span><br><span data-ttu-id="52f90-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-741">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-742">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-742">- Content</span></span><br><span data-ttu-id="52f90-743">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-743">
         - TaskPane</span></span><br><span data-ttu-id="52f90-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-746">- ActiveView</span></span><br><span data-ttu-id="52f90-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-747">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-748">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-749">
         - File</span></span><br><span data-ttu-id="52f90-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-750">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-751">
         - PdfFile</span></span><br><span data-ttu-id="52f90-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-752">
         - Selection</span></span><br><span data-ttu-id="52f90-753">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-753">
         - Settings</span></span><br><span data-ttu-id="52f90-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-755">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-756">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-756">- Content</span></span><br><span data-ttu-id="52f90-757">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-757">
         - TaskPane</span></span><br><span data-ttu-id="52f90-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-760">- ActiveView</span></span><br><span data-ttu-id="52f90-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-761">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-762">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-763">
         - File</span></span><br><span data-ttu-id="52f90-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-764">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-765">
         - PdfFile</span></span><br><span data-ttu-id="52f90-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-766">
         - Selection</span></span><br><span data-ttu-id="52f90-767">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-767">
         - Settings</span></span><br><span data-ttu-id="52f90-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-769">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="52f90-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="52f90-770">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-770">- Content</span></span><br><span data-ttu-id="52f90-771">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-771">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="52f90-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="52f90-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="52f90-773">- ActiveView</span></span><br><span data-ttu-id="52f90-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="52f90-774">
         - CompressedFile</span></span><br><span data-ttu-id="52f90-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-775">
         - DocumentEvents</span></span><br><span data-ttu-id="52f90-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="52f90-776">
         - File</span></span><br><span data-ttu-id="52f90-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-777">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="52f90-778">
         - PdfFile</span></span><br><span data-ttu-id="52f90-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-779">
         - Selection</span></span><br><span data-ttu-id="52f90-780">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-780">
         - Settings</span></span><br><span data-ttu-id="52f90-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="52f90-782">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="52f90-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="52f90-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="52f90-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52f90-784">Платформа</span><span class="sxs-lookup"><span data-stu-id="52f90-784">Platform</span></span></th>
    <th><span data-ttu-id="52f90-785">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="52f90-785">Extension points</span></span></th>
    <th><span data-ttu-id="52f90-786">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="52f90-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="52f90-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="52f90-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="52f90-788">Office Online</span></span></td>
    <td> <span data-ttu-id="52f90-789">- Контент</span><span class="sxs-lookup"><span data-stu-id="52f90-789">- Content</span></span><br><span data-ttu-id="52f90-790">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-790">
         - TaskPane</span></span><br><span data-ttu-id="52f90-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="52f90-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="52f90-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="52f90-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="52f90-794">- DocumentEvents</span></span><br><span data-ttu-id="52f90-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="52f90-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-796">
         - ImageCoercion</span></span><br><span data-ttu-id="52f90-797">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="52f90-797">
         - Settings</span></span><br><span data-ttu-id="52f90-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="52f90-799">Project</span><span class="sxs-lookup"><span data-stu-id="52f90-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="52f90-800">Платформа</span><span class="sxs-lookup"><span data-stu-id="52f90-800">Platform</span></span></th>
    <th><span data-ttu-id="52f90-801">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="52f90-801">Extension points</span></span></th>
    <th><span data-ttu-id="52f90-802">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="52f90-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="52f90-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="52f90-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-804">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-805">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-807">- Selection</span></span><br><span data-ttu-id="52f90-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-809">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-810">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-812">- Selection</span></span><br><span data-ttu-id="52f90-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="52f90-814">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="52f90-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="52f90-815">- Область задач</span><span class="sxs-lookup"><span data-stu-id="52f90-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="52f90-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="52f90-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="52f90-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="52f90-817">- Selection</span></span><br><span data-ttu-id="52f90-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="52f90-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="52f90-819">См. также</span><span class="sxs-lookup"><span data-stu-id="52f90-819">See also</span></span>

- [<span data-ttu-id="52f90-820">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="52f90-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="52f90-821">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="52f90-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="52f90-822">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="52f90-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="52f90-823">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="52f90-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="52f90-824">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="52f90-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="52f90-825">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="52f90-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="52f90-826">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="52f90-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="52f90-827">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="52f90-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)