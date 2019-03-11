---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: 636c6290d8c67901beb195990593727485467460
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512883"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d643d-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d643d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d643d-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="d643d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="d643d-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="d643d-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="d643d-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="d643d-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="d643d-108">Excel</span><span class="sxs-lookup"><span data-stu-id="d643d-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d643d-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="d643d-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d643d-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d643d-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d643d-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d643d-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d643d-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d643d-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="d643d-113">Office Online</span></span></td>
    <td> <span data-ttu-id="d643d-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-114">- TaskPane</span></span><br><span data-ttu-id="d643d-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-115">
        - Content</span></span><br><span data-ttu-id="d643d-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="d643d-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d643d-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d643d-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d643d-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d643d-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d643d-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d643d-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d643d-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d643d-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d643d-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d643d-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d643d-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-126">
        - BindingEvents</span></span><br><span data-ttu-id="d643d-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-127">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-128">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-129">
        - File</span></span><br><span data-ttu-id="d643d-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-130">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-132">
        - Selection</span></span><br><span data-ttu-id="d643d-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-133">
        - Settings</span></span><br><span data-ttu-id="d643d-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-134">
        - TableBindings</span></span><br><span data-ttu-id="d643d-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-135">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-136">
        - TextBindings</span></span><br><span data-ttu-id="d643d-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="d643d-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-139">
        - TaskPane</span></span><br><span data-ttu-id="d643d-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d643d-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d643d-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="d643d-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-142">
        - BindingEvents</span></span><br><span data-ttu-id="d643d-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-143">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-144">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-145">
        - File</span></span><br><span data-ttu-id="d643d-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-146">
        - ImageCoercion</span></span><br><span data-ttu-id="d643d-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-147">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-149">
        - Selection</span></span><br><span data-ttu-id="d643d-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-150">
        - Settings</span></span><br><span data-ttu-id="d643d-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-151">
        - TableBindings</span></span><br><span data-ttu-id="d643d-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-152">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-153">
        - TextBindings</span></span><br><span data-ttu-id="d643d-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="d643d-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-156">- TaskPane</span></span><br><span data-ttu-id="d643d-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-157">
        - Content</span></span></td>
    <td><span data-ttu-id="d643d-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d643d-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d643d-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="d643d-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-160">- BindingEvents</span></span><br><span data-ttu-id="d643d-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-161">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-162">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-163">
        - File</span></span><br><span data-ttu-id="d643d-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-164">
        - ImageCoercion</span></span><br><span data-ttu-id="d643d-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-165">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-167">
        - Selection</span></span><br><span data-ttu-id="d643d-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-168">
        - Settings</span></span><br><span data-ttu-id="d643d-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-169">
        - TableBindings</span></span><br><span data-ttu-id="d643d-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-170">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-171">
        - TextBindings</span></span><br><span data-ttu-id="d643d-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="d643d-174">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-174">- TaskPane</span></span><br><span data-ttu-id="d643d-175">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-175">
        - Content</span></span><br><span data-ttu-id="d643d-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d643d-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d643d-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d643d-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d643d-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d643d-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d643d-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d643d-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d643d-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d643d-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d643d-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d643d-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-186">- BindingEvents</span></span><br><span data-ttu-id="d643d-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-187">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-188">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-189">
        - File</span></span><br><span data-ttu-id="d643d-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-190">
        - ImageCoercion</span></span><br><span data-ttu-id="d643d-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-191">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-193">
        - Selection</span></span><br><span data-ttu-id="d643d-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-194">
        - Settings</span></span><br><span data-ttu-id="d643d-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-195">
        - TableBindings</span></span><br><span data-ttu-id="d643d-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-196">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-197">
        - TextBindings</span></span><br><span data-ttu-id="d643d-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-199">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d643d-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="d643d-200">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-200">- TaskPane</span></span><br><span data-ttu-id="d643d-201">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-201">
        - Content</span></span></td>
    <td><span data-ttu-id="d643d-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d643d-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d643d-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d643d-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d643d-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d643d-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d643d-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d643d-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d643d-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d643d-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d643d-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-211">- BindingEvents</span></span><br><span data-ttu-id="d643d-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-212">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-213">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-214">
        - File</span></span><br><span data-ttu-id="d643d-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-215">
        - ImageCoercion</span></span><br><span data-ttu-id="d643d-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-216">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-218">
        - Selection</span></span><br><span data-ttu-id="d643d-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-219">
        - Settings</span></span><br><span data-ttu-id="d643d-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-220">
        - TableBindings</span></span><br><span data-ttu-id="d643d-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-221">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-222">
        - TextBindings</span></span><br><span data-ttu-id="d643d-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-224">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="d643d-225">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-225">- TaskPane</span></span><br><span data-ttu-id="d643d-226">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-226">
        - Content</span></span></td>
    <td><span data-ttu-id="d643d-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d643d-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d643d-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="d643d-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-229">- BindingEvents</span></span><br><span data-ttu-id="d643d-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-230">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-231">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-232">
        - File</span></span><br><span data-ttu-id="d643d-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-233">
        - ImageCoercion</span></span><br><span data-ttu-id="d643d-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-234">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-236">
        - PdfFile</span></span><br><span data-ttu-id="d643d-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-237">
        - Selection</span></span><br><span data-ttu-id="d643d-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-238">
        - Settings</span></span><br><span data-ttu-id="d643d-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-239">
        - TableBindings</span></span><br><span data-ttu-id="d643d-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-240">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-241">
        - TextBindings</span></span><br><span data-ttu-id="d643d-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-243">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="d643d-244">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-244">- TaskPane</span></span><br><span data-ttu-id="d643d-245">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-245">
        - Content</span></span><br><span data-ttu-id="d643d-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d643d-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d643d-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d643d-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d643d-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d643d-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d643d-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d643d-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="d643d-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d643d-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d643d-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d643d-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-256">- BindingEvents</span></span><br><span data-ttu-id="d643d-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-257">
        - CompressedFile</span></span><br><span data-ttu-id="d643d-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-258">
        - DocumentEvents</span></span><br><span data-ttu-id="d643d-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="d643d-259">
        - File</span></span><br><span data-ttu-id="d643d-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-260">
        - ImageCoercion</span></span><br><span data-ttu-id="d643d-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-261">
        - MatrixBindings</span></span><br><span data-ttu-id="d643d-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="d643d-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-263">
        - PdfFile</span></span><br><span data-ttu-id="d643d-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-264">
        - Selection</span></span><br><span data-ttu-id="d643d-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-265">
        - Settings</span></span><br><span data-ttu-id="d643d-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-266">
        - TableBindings</span></span><br><span data-ttu-id="d643d-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-267">
        - TableCoercion</span></span><br><span data-ttu-id="d643d-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-268">
        - TextBindings</span></span><br><span data-ttu-id="d643d-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d643d-270">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d643d-270">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="d643d-271">Outlook</span><span class="sxs-lookup"><span data-stu-id="d643d-271">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d643d-272">Платформа</span><span class="sxs-lookup"><span data-stu-id="d643d-272">Platform</span></span></th>
    <th><span data-ttu-id="d643d-273">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d643d-273">Extension points</span></span></th>
    <th><span data-ttu-id="d643d-274">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d643d-274">API requirement sets</span></span></th>
    <th><span data-ttu-id="d643d-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d643d-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-276">Office Online</span><span class="sxs-lookup"><span data-stu-id="d643d-276">Office Online</span></span></td>
    <td> <span data-ttu-id="d643d-277">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-277">- Mail Read</span></span><br><span data-ttu-id="d643d-278">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d643d-278">
      - Mail Compose</span></span><br><span data-ttu-id="d643d-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d643d-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d643d-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d643d-287">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-288">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-288">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-289">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-289">- Mail Read</span></span><br><span data-ttu-id="d643d-290">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d643d-290">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="d643d-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="d643d-295">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-295">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-296">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-296">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-297">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-297">- Mail Read</span></span><br><span data-ttu-id="d643d-298">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d643d-298">
      - Mail Compose</span></span><br><span data-ttu-id="d643d-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d643d-300">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d643d-300">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d643d-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d643d-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d643d-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d643d-308">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-308">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-309">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-309">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-310">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-310">- Mail Read</span></span><br><span data-ttu-id="d643d-311">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d643d-311">
      - Mail Compose</span></span><br><span data-ttu-id="d643d-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d643d-313">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="d643d-313">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d643d-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d643d-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d643d-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d643d-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d643d-321">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-321">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-322">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="d643d-322">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d643d-323">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-323">- Mail Read</span></span><br><span data-ttu-id="d643d-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d643d-330">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-330">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-331">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-331">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d643d-332">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-332">- Mail Read</span></span><br><span data-ttu-id="d643d-333">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d643d-333">
      - Mail Compose</span></span><br><span data-ttu-id="d643d-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d643d-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d643d-341">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-341">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-342">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-342">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="d643d-343">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-343">- Mail Read</span></span><br><span data-ttu-id="d643d-344">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="d643d-344">
      - Mail Compose</span></span><br><span data-ttu-id="d643d-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d643d-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d643d-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d643d-352">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-353">Office для Android</span><span class="sxs-lookup"><span data-stu-id="d643d-353">Office for Android</span></span></td>
    <td> <span data-ttu-id="d643d-354">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="d643d-354">- Mail Read</span></span><br><span data-ttu-id="d643d-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d643d-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d643d-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d643d-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d643d-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d643d-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d643d-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d643d-361">Недоступно</span><span class="sxs-lookup"><span data-stu-id="d643d-361">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="d643d-362">Word</span><span class="sxs-lookup"><span data-stu-id="d643d-362">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d643d-363">Платформа</span><span class="sxs-lookup"><span data-stu-id="d643d-363">Platform</span></span></th>
    <th><span data-ttu-id="d643d-364">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d643d-364">Extension points</span></span></th>
    <th><span data-ttu-id="d643d-365">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d643d-365">API requirement sets</span></span></th>
    <th><span data-ttu-id="d643d-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d643d-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-367">Office Online</span><span class="sxs-lookup"><span data-stu-id="d643d-367">Office Online</span></span></td>
    <td> <span data-ttu-id="d643d-368">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-368">- TaskPane</span></span><br><span data-ttu-id="d643d-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d643d-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d643d-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d643d-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-374">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-374">- BindingEvents</span></span><br><span data-ttu-id="d643d-375">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-375">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-376">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-376">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-377">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-377">
         - File</span></span><br><span data-ttu-id="d643d-378">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-378">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-379">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-379">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-380">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-380">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-381">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-381">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-382">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-382">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-383">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-383">
         - PdfFile</span></span><br><span data-ttu-id="d643d-384">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-384">
         - Selection</span></span><br><span data-ttu-id="d643d-385">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-385">
         - Settings</span></span><br><span data-ttu-id="d643d-386">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-386">
         - TableBindings</span></span><br><span data-ttu-id="d643d-387">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-387">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-388">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-388">
         - TextBindings</span></span><br><span data-ttu-id="d643d-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-389">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-390">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-390">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-391">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-392">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-392">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d643d-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="d643d-394">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-394">- BindingEvents</span></span><br><span data-ttu-id="d643d-395">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-395">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-396">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-396">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-397">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-398">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-398">
         - File</span></span><br><span data-ttu-id="d643d-399">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-399">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-400">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-400">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-401">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-401">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-402">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-402">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-403">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-403">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-404">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-404">
         - PdfFile</span></span><br><span data-ttu-id="d643d-405">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-405">
         - Selection</span></span><br><span data-ttu-id="d643d-406">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-406">
         - Settings</span></span><br><span data-ttu-id="d643d-407">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-407">
         - TableBindings</span></span><br><span data-ttu-id="d643d-408">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-408">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-409">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-409">
         - TextBindings</span></span><br><span data-ttu-id="d643d-410">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-410">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-411">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-411">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-412">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-412">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-413">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-413">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d643d-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d643d-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="d643d-416">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-416">- BindingEvents</span></span><br><span data-ttu-id="d643d-417">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-417">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-418">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-418">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-419">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-419">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-420">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-420">
         - File</span></span><br><span data-ttu-id="d643d-421">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-421">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-422">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-422">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-423">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-423">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-424">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-424">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-425">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-425">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-426">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-426">
         - PdfFile</span></span><br><span data-ttu-id="d643d-427">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-427">
         - Selection</span></span><br><span data-ttu-id="d643d-428">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-428">
         - Settings</span></span><br><span data-ttu-id="d643d-429">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-429">
         - TableBindings</span></span><br><span data-ttu-id="d643d-430">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-430">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-431">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-431">
         - TextBindings</span></span><br><span data-ttu-id="d643d-432">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-432">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-433">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-433">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-434">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-434">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-435">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-435">- TaskPane</span></span><br><span data-ttu-id="d643d-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d643d-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d643d-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d643d-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-441">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-441">- BindingEvents</span></span><br><span data-ttu-id="d643d-442">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-442">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-443">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-443">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-444">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-444">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-445">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-445">
         - File</span></span><br><span data-ttu-id="d643d-446">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-446">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-447">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-447">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-448">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-448">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-449">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-449">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-450">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-450">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-451">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-451">
         - PdfFile</span></span><br><span data-ttu-id="d643d-452">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-452">
         - Selection</span></span><br><span data-ttu-id="d643d-453">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-453">
         - Settings</span></span><br><span data-ttu-id="d643d-454">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-454">
         - TableBindings</span></span><br><span data-ttu-id="d643d-455">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-455">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-456">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-456">
         - TextBindings</span></span><br><span data-ttu-id="d643d-457">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-457">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-458">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-458">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-459">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d643d-459">Office for iPad</span></span></td>
    <td> <span data-ttu-id="d643d-460">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-460">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d643d-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d643d-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d643d-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d643d-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d643d-465">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-465">- BindingEvents</span></span><br><span data-ttu-id="d643d-466">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-466">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-467">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-467">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-468">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-468">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-469">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-469">
         - File</span></span><br><span data-ttu-id="d643d-470">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-470">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-471">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-471">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-472">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-472">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-473">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-473">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-474">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-474">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-475">
         - PdfFile</span></span><br><span data-ttu-id="d643d-476">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-476">
         - Selection</span></span><br><span data-ttu-id="d643d-477">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-477">
         - Settings</span></span><br><span data-ttu-id="d643d-478">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-478">
         - TableBindings</span></span><br><span data-ttu-id="d643d-479">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-479">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-480">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-480">
         - TextBindings</span></span><br><span data-ttu-id="d643d-481">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-481">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-482">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-482">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-483">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-483">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d643d-484">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-484">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d643d-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="d643d-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="d643d-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-487">- BindingEvents</span></span><br><span data-ttu-id="d643d-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-488">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-489">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-490">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-491">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-491">
         - File</span></span><br><span data-ttu-id="d643d-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-492">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-493">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-494">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-495">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-496">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-497">
         - PdfFile</span></span><br><span data-ttu-id="d643d-498">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-498">
         - Selection</span></span><br><span data-ttu-id="d643d-499">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-499">
         - Settings</span></span><br><span data-ttu-id="d643d-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-500">
         - TableBindings</span></span><br><span data-ttu-id="d643d-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-501">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-502">
         - TextBindings</span></span><br><span data-ttu-id="d643d-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-503">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-504">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-505">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-505">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="d643d-506">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-506">- TaskPane</span></span><br><span data-ttu-id="d643d-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d643d-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d643d-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d643d-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d643d-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d643d-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d643d-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d643d-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-512">- BindingEvents</span></span><br><span data-ttu-id="d643d-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-513">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d643d-514">
         - CustomXmlParts</span></span><br><span data-ttu-id="d643d-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-515">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-516">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="d643d-516">
         - File</span></span><br><span data-ttu-id="d643d-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-517">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-518">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-519">
         - MatrixBindings</span></span><br><span data-ttu-id="d643d-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-520">
         - MatrixCoercion</span></span><br><span data-ttu-id="d643d-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-521">
         - OoxmlCoercion</span></span><br><span data-ttu-id="d643d-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-522">
         - PdfFile</span></span><br><span data-ttu-id="d643d-523">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-523">
         - Selection</span></span><br><span data-ttu-id="d643d-524">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="d643d-524">
         - Settings</span></span><br><span data-ttu-id="d643d-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-525">
         - TableBindings</span></span><br><span data-ttu-id="d643d-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-526">
         - TableCoercion</span></span><br><span data-ttu-id="d643d-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d643d-527">
         - TextBindings</span></span><br><span data-ttu-id="d643d-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-528">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d643d-529">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="d643d-530">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d643d-530">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d643d-531">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d643d-531">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d643d-532">Платформа</span><span class="sxs-lookup"><span data-stu-id="d643d-532">Platform</span></span></th>
    <th><span data-ttu-id="d643d-533">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d643d-533">Extension points</span></span></th>
    <th><span data-ttu-id="d643d-534">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d643d-534">API requirement sets</span></span></th>
    <th><span data-ttu-id="d643d-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d643d-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-536">Office Online</span><span class="sxs-lookup"><span data-stu-id="d643d-536">Office Online</span></span></td>
    <td> <span data-ttu-id="d643d-537">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-537">- Content</span></span><br><span data-ttu-id="d643d-538">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-538">
         - TaskPane</span></span><br><span data-ttu-id="d643d-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-541">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-541">- ActiveView</span></span><br><span data-ttu-id="d643d-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-542">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-543">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-544">
         - File</span></span><br><span data-ttu-id="d643d-545">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-545">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-546">
         - PdfFile</span></span><br><span data-ttu-id="d643d-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-547">
         - Selection</span></span><br><span data-ttu-id="d643d-548">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-548">
         - Settings</span></span><br><span data-ttu-id="d643d-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-549">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-550">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-550">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-551">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-551">- Content</span></span><br><span data-ttu-id="d643d-552">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-552">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="d643d-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d643d-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="d643d-554">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-554">- ActiveView</span></span><br><span data-ttu-id="d643d-555">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-555">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-556">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-557">
         - File</span></span><br><span data-ttu-id="d643d-558">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-558">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-559">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-559">
         - PdfFile</span></span><br><span data-ttu-id="d643d-560">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-560">
         - Selection</span></span><br><span data-ttu-id="d643d-561">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-561">
         - Settings</span></span><br><span data-ttu-id="d643d-562">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-562">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-563">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-563">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-564">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-564">- Content</span></span><br><span data-ttu-id="d643d-565">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-565">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d643d-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="d643d-567">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-567">- ActiveView</span></span><br><span data-ttu-id="d643d-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-568">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-569">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-570">
         - File</span></span><br><span data-ttu-id="d643d-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-571">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-572">
         - PdfFile</span></span><br><span data-ttu-id="d643d-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-573">
         - Selection</span></span><br><span data-ttu-id="d643d-574">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-574">
         - Settings</span></span><br><span data-ttu-id="d643d-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-575">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-576">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-576">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-577">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-577">- Content</span></span><br><span data-ttu-id="d643d-578">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-578">
         - TaskPane</span></span><br><span data-ttu-id="d643d-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-581">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-581">- ActiveView</span></span><br><span data-ttu-id="d643d-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-582">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-583">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-583">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-584">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-584">
         - File</span></span><br><span data-ttu-id="d643d-585">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-585">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-586">
         - PdfFile</span></span><br><span data-ttu-id="d643d-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-587">
         - Selection</span></span><br><span data-ttu-id="d643d-588">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-588">
         - Settings</span></span><br><span data-ttu-id="d643d-589">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-589">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-590">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="d643d-590">Office for iPad</span></span></td>
    <td> <span data-ttu-id="d643d-591">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-591">- Content</span></span><br><span data-ttu-id="d643d-592">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-592">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="d643d-594">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-594">- ActiveView</span></span><br><span data-ttu-id="d643d-595">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-595">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-596">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-596">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-597">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-597">
         - File</span></span><br><span data-ttu-id="d643d-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-598">
         - PdfFile</span></span><br><span data-ttu-id="d643d-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-599">
         - Selection</span></span><br><span data-ttu-id="d643d-600">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-600">
         - Settings</span></span><br><span data-ttu-id="d643d-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-601">
         - TextCoercion</span></span><br><span data-ttu-id="d643d-602">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-602">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-603">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-603">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d643d-604">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-604">- Content</span></span><br><span data-ttu-id="d643d-605">
         - Область задач/td></span><span class="sxs-lookup"><span data-stu-id="d643d-605">
         - TaskPane/td></span></span> <td> <span data-ttu-id="d643d-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="d643d-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="d643d-607">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-607">- ActiveView</span></span><br><span data-ttu-id="d643d-608">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-608">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-609">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-610">
         - File</span></span><br><span data-ttu-id="d643d-611">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-611">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-612">
         - PdfFile</span></span><br><span data-ttu-id="d643d-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-613">
         - Selection</span></span><br><span data-ttu-id="d643d-614">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-614">
         - Settings</span></span><br><span data-ttu-id="d643d-615">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-615">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-616">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="d643d-616">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="d643d-617">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-617">- Content</span></span><br><span data-ttu-id="d643d-618">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-618">
         - TaskPane</span></span><br><span data-ttu-id="d643d-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-621">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d643d-621">- ActiveView</span></span><br><span data-ttu-id="d643d-622">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d643d-622">
         - CompressedFile</span></span><br><span data-ttu-id="d643d-623">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-623">
         - DocumentEvents</span></span><br><span data-ttu-id="d643d-624">
         - File</span><span class="sxs-lookup"><span data-stu-id="d643d-624">
         - File</span></span><br><span data-ttu-id="d643d-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-625">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-626">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d643d-626">
         - PdfFile</span></span><br><span data-ttu-id="d643d-627">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-627">
         - Selection</span></span><br><span data-ttu-id="d643d-628">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-628">
         - Settings</span></span><br><span data-ttu-id="d643d-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-629">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="d643d-630">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="d643d-630">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="d643d-631">OneNote</span><span class="sxs-lookup"><span data-stu-id="d643d-631">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d643d-632">Платформа</span><span class="sxs-lookup"><span data-stu-id="d643d-632">Platform</span></span></th>
    <th><span data-ttu-id="d643d-633">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d643d-633">Extension points</span></span></th>
    <th><span data-ttu-id="d643d-634">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d643d-634">API requirement sets</span></span></th>
    <th><span data-ttu-id="d643d-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d643d-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-636">Office Online</span><span class="sxs-lookup"><span data-stu-id="d643d-636">Office Online</span></span></td>
    <td> <span data-ttu-id="d643d-637">- Контент</span><span class="sxs-lookup"><span data-stu-id="d643d-637">- Content</span></span><br><span data-ttu-id="d643d-638">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-638">
         - TaskPane</span></span><br><span data-ttu-id="d643d-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="d643d-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d643d-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d643d-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-642">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d643d-642">- DocumentEvents</span></span><br><span data-ttu-id="d643d-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="d643d-644">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-644">
         - ImageCoercion</span></span><br><span data-ttu-id="d643d-645">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="d643d-645">
         - Settings</span></span><br><span data-ttu-id="d643d-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-646">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="d643d-647">Project</span><span class="sxs-lookup"><span data-stu-id="d643d-647">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d643d-648">Платформа</span><span class="sxs-lookup"><span data-stu-id="d643d-648">Platform</span></span></th>
    <th><span data-ttu-id="d643d-649">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="d643d-649">Extension points</span></span></th>
    <th><span data-ttu-id="d643d-650">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="d643d-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="d643d-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="d643d-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-652">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-652">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-653">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-653">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-655">- Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-655">- Selection</span></span><br><span data-ttu-id="d643d-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-656">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-657">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-657">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-658">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-658">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-660">- Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-660">- Selection</span></span><br><span data-ttu-id="d643d-661">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-661">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d643d-662">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="d643d-662">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="d643d-663">- Область задач</span><span class="sxs-lookup"><span data-stu-id="d643d-663">- TaskPane</span></span></td>
    <td> <span data-ttu-id="d643d-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d643d-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d643d-665">- Selection</span><span class="sxs-lookup"><span data-stu-id="d643d-665">- Selection</span></span><br><span data-ttu-id="d643d-666">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d643d-666">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d643d-667">См. также</span><span class="sxs-lookup"><span data-stu-id="d643d-667">See also</span></span>

- [<span data-ttu-id="d643d-668">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d643d-668">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d643d-669">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="d643d-669">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="d643d-670">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="d643d-670">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="d643d-671">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="d643d-671">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
