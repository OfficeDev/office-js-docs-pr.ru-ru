---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 28a6d0e4c86d05855ed9d24461dbeb77454d2b48
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872132"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c527e-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c527e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c527e-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="c527e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="c527e-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="c527e-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c527e-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="c527e-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="c527e-108">Номер сборки для единовременной покупки Office 2019 — 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="c527e-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="c527e-109">Excel</span><span class="sxs-lookup"><span data-stu-id="c527e-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c527e-110">Платформа</span><span class="sxs-lookup"><span data-stu-id="c527e-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c527e-111">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c527e-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c527e-112">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c527e-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c527e-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c527e-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="c527e-114">Office Online</span></span></td>
    <td> <span data-ttu-id="c527e-115">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-115">- TaskPane</span></span><br><span data-ttu-id="c527e-116">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-116">
        - Content</span></span><br><span data-ttu-id="c527e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="c527e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c527e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c527e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c527e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c527e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c527e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c527e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c527e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c527e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c527e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c527e-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-127">
        - BindingEvents</span></span><br><span data-ttu-id="c527e-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-128">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-129">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-130">
        - File</span></span><br><span data-ttu-id="c527e-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-131">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-133">
        - Selection</span></span><br><span data-ttu-id="c527e-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-134">
        - Settings</span></span><br><span data-ttu-id="c527e-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-135">
        - TableBindings</span></span><br><span data-ttu-id="c527e-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-136">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-137">
        - TextBindings</span></span><br><span data-ttu-id="c527e-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-139">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-140">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-140">- TaskPane</span></span><br><span data-ttu-id="c527e-141">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-141">
        - Content</span></span><br><span data-ttu-id="c527e-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="c527e-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c527e-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c527e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c527e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c527e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c527e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c527e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c527e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c527e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c527e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c527e-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-152">
        - BindingEvents</span></span><br><span data-ttu-id="c527e-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-153">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-154">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-155">
        - File</span></span><br><span data-ttu-id="c527e-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-156">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-158">
        - Selection</span></span><br><span data-ttu-id="c527e-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-159">
        - Settings</span></span><br><span data-ttu-id="c527e-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-160">
        - TableBindings</span></span><br><span data-ttu-id="c527e-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-161">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-162">
        - TextBindings</span></span><br><span data-ttu-id="c527e-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-164">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="c527e-165">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-165">- TaskPane</span></span><br><span data-ttu-id="c527e-166">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-166">
        - Content</span></span><br><span data-ttu-id="c527e-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c527e-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c527e-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c527e-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c527e-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c527e-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c527e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c527e-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c527e-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c527e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c527e-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-177">- BindingEvents</span></span><br><span data-ttu-id="c527e-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-178">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-179">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-180">
        - File</span></span><br><span data-ttu-id="c527e-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-181">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-182">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-184">
        - Selection</span></span><br><span data-ttu-id="c527e-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-185">
        - Settings</span></span><br><span data-ttu-id="c527e-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-186">
        - TableBindings</span></span><br><span data-ttu-id="c527e-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-187">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-188">
        - TextBindings</span></span><br><span data-ttu-id="c527e-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-190">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="c527e-191">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-191">- TaskPane</span></span><br><span data-ttu-id="c527e-192">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-192">
        - Content</span></span></td>
    <td><span data-ttu-id="c527e-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="c527e-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-195">- BindingEvents</span></span><br><span data-ttu-id="c527e-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-196">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-197">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-198">
        - File</span></span><br><span data-ttu-id="c527e-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-199">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-200">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-202">
        - Selection</span></span><br><span data-ttu-id="c527e-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-203">
        - Settings</span></span><br><span data-ttu-id="c527e-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-204">
        - TableBindings</span></span><br><span data-ttu-id="c527e-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-205">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-206">
        - TextBindings</span></span><br><span data-ttu-id="c527e-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-208">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="c527e-209">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-209">
        - TaskPane</span></span><br><span data-ttu-id="c527e-210">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c527e-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c527e-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="c527e-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-212">
        - BindingEvents</span></span><br><span data-ttu-id="c527e-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-213">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-214">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-215">
        - File</span></span><br><span data-ttu-id="c527e-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-216">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-217">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-219">
        - Selection</span></span><br><span data-ttu-id="c527e-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-220">
        - Settings</span></span><br><span data-ttu-id="c527e-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-221">
        - TableBindings</span></span><br><span data-ttu-id="c527e-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-222">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-223">
        - TextBindings</span></span><br><span data-ttu-id="c527e-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-225">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="c527e-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="c527e-226">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-226">- TaskPane</span></span><br><span data-ttu-id="c527e-227">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-227">
        - Content</span></span></td>
    <td><span data-ttu-id="c527e-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c527e-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c527e-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c527e-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c527e-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c527e-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c527e-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c527e-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c527e-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c527e-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-237">- BindingEvents</span></span><br><span data-ttu-id="c527e-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-238">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-239">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-240">
        - File</span></span><br><span data-ttu-id="c527e-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-241">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-242">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-244">
        - Selection</span></span><br><span data-ttu-id="c527e-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-245">
        - Settings</span></span><br><span data-ttu-id="c527e-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-246">
        - TableBindings</span></span><br><span data-ttu-id="c527e-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-247">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-248">
        - TextBindings</span></span><br><span data-ttu-id="c527e-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-250">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="c527e-251">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-251">- TaskPane</span></span><br><span data-ttu-id="c527e-252">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-252">
        - Content</span></span><br><span data-ttu-id="c527e-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c527e-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c527e-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c527e-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c527e-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c527e-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c527e-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c527e-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c527e-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c527e-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c527e-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-263">- BindingEvents</span></span><br><span data-ttu-id="c527e-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-264">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-265">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-266">
        - File</span></span><br><span data-ttu-id="c527e-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-267">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-268">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-270">
        - PdfFile</span></span><br><span data-ttu-id="c527e-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-271">
        - Selection</span></span><br><span data-ttu-id="c527e-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-272">
        - Settings</span></span><br><span data-ttu-id="c527e-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-273">
        - TableBindings</span></span><br><span data-ttu-id="c527e-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-274">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-275">
        - TextBindings</span></span><br><span data-ttu-id="c527e-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-277">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="c527e-278">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-278">- TaskPane</span></span><br><span data-ttu-id="c527e-279">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-279">
        - Content</span></span><br><span data-ttu-id="c527e-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c527e-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c527e-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c527e-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c527e-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c527e-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c527e-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c527e-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c527e-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c527e-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c527e-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-290">- BindingEvents</span></span><br><span data-ttu-id="c527e-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-291">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-292">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-293">
        - File</span></span><br><span data-ttu-id="c527e-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-294">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-295">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-297">
        - PdfFile</span></span><br><span data-ttu-id="c527e-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-298">
        - Selection</span></span><br><span data-ttu-id="c527e-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-299">
        - Settings</span></span><br><span data-ttu-id="c527e-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-300">
        - TableBindings</span></span><br><span data-ttu-id="c527e-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-301">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-302">
        - TextBindings</span></span><br><span data-ttu-id="c527e-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-304">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="c527e-305">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-305">- TaskPane</span></span><br><span data-ttu-id="c527e-306">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-306">
        - Content</span></span></td>
    <td><span data-ttu-id="c527e-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c527e-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="c527e-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-309">- BindingEvents</span></span><br><span data-ttu-id="c527e-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-310">
        - CompressedFile</span></span><br><span data-ttu-id="c527e-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-311">
        - DocumentEvents</span></span><br><span data-ttu-id="c527e-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="c527e-312">
        - File</span></span><br><span data-ttu-id="c527e-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-313">
        - ImageCoercion</span></span><br><span data-ttu-id="c527e-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-314">
        - MatrixBindings</span></span><br><span data-ttu-id="c527e-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="c527e-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-316">
        - PdfFile</span></span><br><span data-ttu-id="c527e-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-317">
        - Selection</span></span><br><span data-ttu-id="c527e-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-318">
        - Settings</span></span><br><span data-ttu-id="c527e-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-319">
        - TableBindings</span></span><br><span data-ttu-id="c527e-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-320">
        - TableCoercion</span></span><br><span data-ttu-id="c527e-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-321">
        - TextBindings</span></span><br><span data-ttu-id="c527e-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c527e-323">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c527e-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="c527e-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="c527e-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c527e-325">Платформа</span><span class="sxs-lookup"><span data-stu-id="c527e-325">Platform</span></span></th>
    <th><span data-ttu-id="c527e-326">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c527e-326">Extension points</span></span></th>
    <th><span data-ttu-id="c527e-327">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c527e-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="c527e-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c527e-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="c527e-329">Office Online</span></span></td>
    <td> <span data-ttu-id="c527e-330">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-330">- Mail Read</span></span><br><span data-ttu-id="c527e-331">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-331">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c527e-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c527e-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c527e-340">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-341">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-342">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-342">- Mail Read</span></span><br><span data-ttu-id="c527e-343">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-343">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c527e-345">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="c527e-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c527e-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c527e-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c527e-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c527e-353">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-354">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-355">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-355">- Mail Read</span></span><br><span data-ttu-id="c527e-356">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-356">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c527e-358">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="c527e-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c527e-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c527e-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c527e-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c527e-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c527e-366">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-367">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-368">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-368">- Mail Read</span></span><br><span data-ttu-id="c527e-369">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-369">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c527e-371">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="c527e-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c527e-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c527e-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-377">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-378">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-378">- Mail Read</span></span><br><span data-ttu-id="c527e-379">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="c527e-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c527e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c527e-384">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-385">Office 365 для iOS</span><span class="sxs-lookup"><span data-stu-id="c527e-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="c527e-386">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-386">- Mail Read</span></span><br><span data-ttu-id="c527e-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c527e-393">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-394">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-395">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-395">- Mail Read</span></span><br><span data-ttu-id="c527e-396">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-396">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c527e-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c527e-404">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-405">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-406">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-406">- Mail Read</span></span><br><span data-ttu-id="c527e-407">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-407">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c527e-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c527e-415">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-416">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-417">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-417">- Mail Read</span></span><br><span data-ttu-id="c527e-418">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c527e-418">
      - Mail Compose</span></span><br><span data-ttu-id="c527e-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c527e-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c527e-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c527e-426">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-427">Office 365 для Android</span><span class="sxs-lookup"><span data-stu-id="c527e-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="c527e-428">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c527e-428">- Mail Read</span></span><br><span data-ttu-id="c527e-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c527e-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c527e-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c527e-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c527e-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c527e-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c527e-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c527e-435">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c527e-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c527e-436">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c527e-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c527e-437">Word</span><span class="sxs-lookup"><span data-stu-id="c527e-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c527e-438">Платформа</span><span class="sxs-lookup"><span data-stu-id="c527e-438">Platform</span></span></th>
    <th><span data-ttu-id="c527e-439">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c527e-439">Extension points</span></span></th>
    <th><span data-ttu-id="c527e-440">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c527e-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="c527e-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c527e-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="c527e-442">Office Online</span></span></td>
    <td> <span data-ttu-id="c527e-443">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-443">- TaskPane</span></span><br><span data-ttu-id="c527e-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c527e-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c527e-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-449">- BindingEvents</span></span><br><span data-ttu-id="c527e-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-451">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-452">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-452">
         - File</span></span><br><span data-ttu-id="c527e-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-454">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-455">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-458">
         - PdfFile</span></span><br><span data-ttu-id="c527e-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-459">
         - Selection</span></span><br><span data-ttu-id="c527e-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-460">
         - Settings</span></span><br><span data-ttu-id="c527e-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-461">
         - TableBindings</span></span><br><span data-ttu-id="c527e-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-462">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-463">
         - TextBindings</span></span><br><span data-ttu-id="c527e-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-464">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-466">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-467">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-467">- TaskPane</span></span><br><span data-ttu-id="c527e-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c527e-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c527e-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-473">- BindingEvents</span></span><br><span data-ttu-id="c527e-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-474">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-476">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-477">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-477">
         - File</span></span><br><span data-ttu-id="c527e-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-479">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-480">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-483">
         - PdfFile</span></span><br><span data-ttu-id="c527e-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-484">
         - Selection</span></span><br><span data-ttu-id="c527e-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-485">
         - Settings</span></span><br><span data-ttu-id="c527e-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-486">
         - TableBindings</span></span><br><span data-ttu-id="c527e-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-487">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-488">
         - TextBindings</span></span><br><span data-ttu-id="c527e-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-489">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-491">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-492">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-492">- TaskPane</span></span><br><span data-ttu-id="c527e-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c527e-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c527e-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-498">- BindingEvents</span></span><br><span data-ttu-id="c527e-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-499">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-501">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-502">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-502">
         - File</span></span><br><span data-ttu-id="c527e-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-504">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-505">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-508">
         - PdfFile</span></span><br><span data-ttu-id="c527e-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-509">
         - Selection</span></span><br><span data-ttu-id="c527e-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-510">
         - Settings</span></span><br><span data-ttu-id="c527e-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-511">
         - TableBindings</span></span><br><span data-ttu-id="c527e-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-512">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-513">
         - TextBindings</span></span><br><span data-ttu-id="c527e-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-514">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-516">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-517">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="c527e-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-520">- BindingEvents</span></span><br><span data-ttu-id="c527e-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-521">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-523">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-524">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-524">
         - File</span></span><br><span data-ttu-id="c527e-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-526">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-527">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-530">
         - PdfFile</span></span><br><span data-ttu-id="c527e-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-531">
         - Selection</span></span><br><span data-ttu-id="c527e-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-532">
         - Settings</span></span><br><span data-ttu-id="c527e-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-533">
         - TableBindings</span></span><br><span data-ttu-id="c527e-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-534">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-535">
         - TextBindings</span></span><br><span data-ttu-id="c527e-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-536">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-538">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-539">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c527e-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="c527e-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-541">- BindingEvents</span></span><br><span data-ttu-id="c527e-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-542">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-544">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-545">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-545">
         - File</span></span><br><span data-ttu-id="c527e-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-547">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-548">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-551">
         - PdfFile</span></span><br><span data-ttu-id="c527e-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-552">
         - Selection</span></span><br><span data-ttu-id="c527e-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-553">
         - Settings</span></span><br><span data-ttu-id="c527e-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-554">
         - TableBindings</span></span><br><span data-ttu-id="c527e-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-555">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-556">
         - TextBindings</span></span><br><span data-ttu-id="c527e-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-557">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-559">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="c527e-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="c527e-560">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c527e-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c527e-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c527e-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c527e-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-565">- BindingEvents</span></span><br><span data-ttu-id="c527e-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-566">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-568">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-569">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-569">
         - File</span></span><br><span data-ttu-id="c527e-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-571">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-572">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-575">
         - PdfFile</span></span><br><span data-ttu-id="c527e-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-576">
         - Selection</span></span><br><span data-ttu-id="c527e-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-577">
         - Settings</span></span><br><span data-ttu-id="c527e-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-578">
         - TableBindings</span></span><br><span data-ttu-id="c527e-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-579">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-580">
         - TextBindings</span></span><br><span data-ttu-id="c527e-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-581">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-583">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-584">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-584">- TaskPane</span></span><br><span data-ttu-id="c527e-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c527e-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c527e-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c527e-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c527e-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-590">- BindingEvents</span></span><br><span data-ttu-id="c527e-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-591">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-593">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-594">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-594">
         - File</span></span><br><span data-ttu-id="c527e-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-596">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-600">
         - PdfFile</span></span><br><span data-ttu-id="c527e-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-601">
         - Selection</span></span><br><span data-ttu-id="c527e-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-602">
         - Settings</span></span><br><span data-ttu-id="c527e-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-603">
         - TableBindings</span></span><br><span data-ttu-id="c527e-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-604">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-605">
         - TextBindings</span></span><br><span data-ttu-id="c527e-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-606">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-608">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-609">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-609">- TaskPane</span></span><br><span data-ttu-id="c527e-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c527e-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c527e-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c527e-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c527e-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c527e-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c527e-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-615">- BindingEvents</span></span><br><span data-ttu-id="c527e-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-616">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-618">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-619">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-619">
         - File</span></span><br><span data-ttu-id="c527e-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-621">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-622">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-625">
         - PdfFile</span></span><br><span data-ttu-id="c527e-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-626">
         - Selection</span></span><br><span data-ttu-id="c527e-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-627">
         - Settings</span></span><br><span data-ttu-id="c527e-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-628">
         - TableBindings</span></span><br><span data-ttu-id="c527e-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-629">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-630">
         - TextBindings</span></span><br><span data-ttu-id="c527e-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-631">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-633">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-634">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c527e-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c527e-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="c527e-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-637">- BindingEvents</span></span><br><span data-ttu-id="c527e-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-638">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c527e-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="c527e-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-640">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-641">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c527e-641">
         - File</span></span><br><span data-ttu-id="c527e-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-643">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-644">
         - MatrixBindings</span></span><br><span data-ttu-id="c527e-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="c527e-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c527e-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-647">
         - PdfFile</span></span><br><span data-ttu-id="c527e-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-648">
         - Selection</span></span><br><span data-ttu-id="c527e-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c527e-649">
         - Settings</span></span><br><span data-ttu-id="c527e-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-650">
         - TableBindings</span></span><br><span data-ttu-id="c527e-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-651">
         - TableCoercion</span></span><br><span data-ttu-id="c527e-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c527e-652">
         - TextBindings</span></span><br><span data-ttu-id="c527e-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-653">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c527e-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c527e-655">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c527e-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c527e-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c527e-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c527e-657">Платформа</span><span class="sxs-lookup"><span data-stu-id="c527e-657">Platform</span></span></th>
    <th><span data-ttu-id="c527e-658">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c527e-658">Extension points</span></span></th>
    <th><span data-ttu-id="c527e-659">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c527e-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="c527e-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c527e-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="c527e-661">Office Online</span></span></td>
    <td> <span data-ttu-id="c527e-662">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-662">- Content</span></span><br><span data-ttu-id="c527e-663">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-663">
         - TaskPane</span></span><br><span data-ttu-id="c527e-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-666">- ActiveView</span></span><br><span data-ttu-id="c527e-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-667">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-668">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-669">
         - File</span></span><br><span data-ttu-id="c527e-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-670">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-671">
         - PdfFile</span></span><br><span data-ttu-id="c527e-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-672">
         - Selection</span></span><br><span data-ttu-id="c527e-673">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-673">
         - Settings</span></span><br><span data-ttu-id="c527e-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-675">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-676">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-676">- Content</span></span><br><span data-ttu-id="c527e-677">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-677">
         - TaskPane</span></span><br><span data-ttu-id="c527e-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-680">- ActiveView</span></span><br><span data-ttu-id="c527e-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-681">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-682">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-683">
         - File</span></span><br><span data-ttu-id="c527e-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-684">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-685">
         - PdfFile</span></span><br><span data-ttu-id="c527e-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-686">
         - Selection</span></span><br><span data-ttu-id="c527e-687">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-687">
         - Settings</span></span><br><span data-ttu-id="c527e-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-689">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-690">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-690">- Content</span></span><br><span data-ttu-id="c527e-691">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-691">
         - TaskPane</span></span><br><span data-ttu-id="c527e-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-694">- ActiveView</span></span><br><span data-ttu-id="c527e-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-695">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-696">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-697">
         - File</span></span><br><span data-ttu-id="c527e-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-698">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-699">
         - PdfFile</span></span><br><span data-ttu-id="c527e-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-700">
         - Selection</span></span><br><span data-ttu-id="c527e-701">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-701">
         - Settings</span></span><br><span data-ttu-id="c527e-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-703">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-704">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-704">- Content</span></span><br><span data-ttu-id="c527e-705">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c527e-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="c527e-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-707">- ActiveView</span></span><br><span data-ttu-id="c527e-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-708">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-709">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-710">
         - File</span></span><br><span data-ttu-id="c527e-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-711">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-712">
         - PdfFile</span></span><br><span data-ttu-id="c527e-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-713">
         - Selection</span></span><br><span data-ttu-id="c527e-714">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-714">
         - Settings</span></span><br><span data-ttu-id="c527e-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-716">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-717">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-717">- Content</span></span><br><span data-ttu-id="c527e-718">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c527e-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c527e-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="c527e-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-720">- ActiveView</span></span><br><span data-ttu-id="c527e-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-721">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-722">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-723">
         - File</span></span><br><span data-ttu-id="c527e-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-724">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-725">
         - PdfFile</span></span><br><span data-ttu-id="c527e-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-726">
         - Selection</span></span><br><span data-ttu-id="c527e-727">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-727">
         - Settings</span></span><br><span data-ttu-id="c527e-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-729">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="c527e-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="c527e-730">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-730">- Content</span></span><br><span data-ttu-id="c527e-731">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="c527e-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-733">- ActiveView</span></span><br><span data-ttu-id="c527e-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-734">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-735">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-736">
         - File</span></span><br><span data-ttu-id="c527e-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-737">
         - PdfFile</span></span><br><span data-ttu-id="c527e-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-738">
         - Selection</span></span><br><span data-ttu-id="c527e-739">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-739">
         - Settings</span></span><br><span data-ttu-id="c527e-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-740">
         - TextCoercion</span></span><br><span data-ttu-id="c527e-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-742">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-743">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-743">- Content</span></span><br><span data-ttu-id="c527e-744">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-744">
         - TaskPane</span></span><br><span data-ttu-id="c527e-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-747">- ActiveView</span></span><br><span data-ttu-id="c527e-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-748">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-749">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-750">
         - File</span></span><br><span data-ttu-id="c527e-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-751">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-752">
         - PdfFile</span></span><br><span data-ttu-id="c527e-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-753">
         - Selection</span></span><br><span data-ttu-id="c527e-754">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-754">
         - Settings</span></span><br><span data-ttu-id="c527e-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-756">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-757">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-757">- Content</span></span><br><span data-ttu-id="c527e-758">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-758">
         - TaskPane</span></span><br><span data-ttu-id="c527e-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-761">- ActiveView</span></span><br><span data-ttu-id="c527e-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-762">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-763">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-764">
         - File</span></span><br><span data-ttu-id="c527e-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-765">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-766">
         - PdfFile</span></span><br><span data-ttu-id="c527e-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-767">
         - Selection</span></span><br><span data-ttu-id="c527e-768">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-768">
         - Settings</span></span><br><span data-ttu-id="c527e-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-770">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c527e-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c527e-771">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-771">- Content</span></span><br><span data-ttu-id="c527e-772">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c527e-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="c527e-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c527e-774">- ActiveView</span></span><br><span data-ttu-id="c527e-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c527e-775">
         - CompressedFile</span></span><br><span data-ttu-id="c527e-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-776">
         - DocumentEvents</span></span><br><span data-ttu-id="c527e-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="c527e-777">
         - File</span></span><br><span data-ttu-id="c527e-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-778">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c527e-779">
         - PdfFile</span></span><br><span data-ttu-id="c527e-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-780">
         - Selection</span></span><br><span data-ttu-id="c527e-781">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-781">
         - Settings</span></span><br><span data-ttu-id="c527e-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c527e-783">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c527e-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c527e-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="c527e-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c527e-785">Платформа</span><span class="sxs-lookup"><span data-stu-id="c527e-785">Platform</span></span></th>
    <th><span data-ttu-id="c527e-786">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c527e-786">Extension points</span></span></th>
    <th><span data-ttu-id="c527e-787">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c527e-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="c527e-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c527e-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="c527e-789">Office Online</span></span></td>
    <td> <span data-ttu-id="c527e-790">- Контент</span><span class="sxs-lookup"><span data-stu-id="c527e-790">- Content</span></span><br><span data-ttu-id="c527e-791">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-791">
         - TaskPane</span></span><br><span data-ttu-id="c527e-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c527e-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c527e-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c527e-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c527e-795">- DocumentEvents</span></span><br><span data-ttu-id="c527e-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="c527e-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-797">
         - ImageCoercion</span></span><br><span data-ttu-id="c527e-798">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c527e-798">
         - Settings</span></span><br><span data-ttu-id="c527e-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c527e-800">Project</span><span class="sxs-lookup"><span data-stu-id="c527e-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c527e-801">Платформа</span><span class="sxs-lookup"><span data-stu-id="c527e-801">Platform</span></span></th>
    <th><span data-ttu-id="c527e-802">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c527e-802">Extension points</span></span></th>
    <th><span data-ttu-id="c527e-803">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c527e-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="c527e-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c527e-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-805">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-806">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-808">- Selection</span></span><br><span data-ttu-id="c527e-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-810">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-811">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-813">- Selection</span></span><br><span data-ttu-id="c527e-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c527e-815">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c527e-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c527e-816">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c527e-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c527e-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c527e-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c527e-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="c527e-818">- Selection</span></span><br><span data-ttu-id="c527e-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c527e-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c527e-820">См. также</span><span class="sxs-lookup"><span data-stu-id="c527e-820">See also</span></span>

- [<span data-ttu-id="c527e-821">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c527e-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c527e-822">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="c527e-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="c527e-823">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="c527e-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="c527e-824">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="c527e-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
