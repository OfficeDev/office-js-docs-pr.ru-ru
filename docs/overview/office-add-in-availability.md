---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 39a80f322c282e29e6e8c4363f0c82522b33b75d
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579928"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="06025-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="06025-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="06025-p101">Работа надстройки Office должным образом может зависеть от ведущего приложения Office, набора требований, элемента или версии API. В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="06025-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="06025-p102">Если ячейка таблицы содержит символ звездочки (\*), это означает, что поддержка скоро появится. С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="06025-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="06025-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="06025-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="06025-110">Excel</span><span class="sxs-lookup"><span data-stu-id="06025-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="06025-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="06025-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="06025-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="06025-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="06025-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="06025-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="06025-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="06025-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="06025-115">Office Online</span></span></td>
    <td> <span data-ttu-id="06025-116">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-116">- Taskpane</span></span><br><span data-ttu-id="06025-117">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-117">
        - Content</span></span><br><span data-ttu-id="06025-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a>
    </span><span class="sxs-lookup"><span data-stu-id="06025-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="06025-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="06025-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="06025-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="06025-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="06025-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="06025-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="06025-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="06025-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="06025-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="06025-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-128">
        -BindingEvents</span></span><br><span data-ttu-id="06025-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-129">
        -CompressedFile</span></span><br><span data-ttu-id="06025-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-130">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-131">
        - File</span></span><br><span data-ttu-id="06025-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-132">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-134">
        - Selection</span></span><br><span data-ttu-id="06025-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-135">
        - Settings</span></span><br><span data-ttu-id="06025-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-136">
        -TableBindings</span></span><br><span data-ttu-id="06025-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-137">
        -TableCoercion</span></span><br><span data-ttu-id="06025-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-138">
        -TextBindings</span></span><br><span data-ttu-id="06025-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-140">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="06025-141">
        - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-141">
        - Taskpane</span></span><br><span data-ttu-id="06025-142">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="06025-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-144">
        -BindingEvents</span></span><br><span data-ttu-id="06025-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-145">
        -CompressedFile</span></span><br><span data-ttu-id="06025-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-146">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-147">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-147">
        - File</span></span><br><span data-ttu-id="06025-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-148">
        -ImageCoercion</span></span><br><span data-ttu-id="06025-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-149">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-151">
        - Selection</span></span><br><span data-ttu-id="06025-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-152">
        - Settings</span></span><br><span data-ttu-id="06025-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-153">
        -TableBindings</span></span><br><span data-ttu-id="06025-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-154">
        -TableCoercion</span></span><br><span data-ttu-id="06025-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-155">
        -TextBindings</span></span><br><span data-ttu-id="06025-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-157">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="06025-158">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-158">- Taskpane</span></span><br><span data-ttu-id="06025-159">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-159">
        - Content</span></span><br><span data-ttu-id="06025-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="06025-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="06025-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="06025-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="06025-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="06025-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="06025-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="06025-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="06025-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="06025-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="06025-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-170">-BindingEvents</span></span><br><span data-ttu-id="06025-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-171">
        -CompressedFile</span></span><br><span data-ttu-id="06025-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-172">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-173">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-173">
        - File</span></span><br><span data-ttu-id="06025-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-174">
        -ImageCoercion</span></span><br><span data-ttu-id="06025-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-175">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-177">
        - Selection</span></span><br><span data-ttu-id="06025-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-178">
        - Settings</span></span><br><span data-ttu-id="06025-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-179">
        -TableBindings</span></span><br><span data-ttu-id="06025-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-180">
        -TableCoercion</span></span><br><span data-ttu-id="06025-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-181">
        -TextBindings</span></span><br><span data-ttu-id="06025-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-183">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="06025-184">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-184">- Taskpane</span></span><br><span data-ttu-id="06025-185">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-185">
        - Content</span></span><br><span data-ttu-id="06025-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="06025-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="06025-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="06025-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="06025-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="06025-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="06025-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="06025-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="06025-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="06025-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="06025-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-196">-BindingEvents</span></span><br><span data-ttu-id="06025-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-197">
        -CompressedFile</span></span><br><span data-ttu-id="06025-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-198">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-199">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-199">
        - File</span></span><br><span data-ttu-id="06025-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-200">
        -ImageCoercion</span></span><br><span data-ttu-id="06025-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-201">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-203">
        - Selection</span></span><br><span data-ttu-id="06025-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-204">
        - Settings</span></span><br><span data-ttu-id="06025-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-205">
        -TableBindings</span></span><br><span data-ttu-id="06025-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-206">
        -TableCoercion</span></span><br><span data-ttu-id="06025-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-207">
        -TextBindings</span></span><br><span data-ttu-id="06025-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-209">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="06025-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="06025-210">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-210">- Taskpane</span></span><br><span data-ttu-id="06025-211">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-211">
        - Content</span></span></td>
    <td><span data-ttu-id="06025-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="06025-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="06025-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="06025-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="06025-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="06025-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="06025-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="06025-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="06025-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="06025-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-221">-BindingEvents</span></span><br><span data-ttu-id="06025-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-222">
        -CompressedFile</span></span><br><span data-ttu-id="06025-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-223">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-224">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-224">
        - File</span></span><br><span data-ttu-id="06025-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-225">
        -ImageCoercion</span></span><br><span data-ttu-id="06025-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-226">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-228">
        - Selection</span></span><br><span data-ttu-id="06025-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-229">
        - Settings</span></span><br><span data-ttu-id="06025-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-230">
        -TableBindings</span></span><br><span data-ttu-id="06025-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-231">
        -TableCoercion</span></span><br><span data-ttu-id="06025-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-232">
        -TextBindings</span></span><br><span data-ttu-id="06025-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-234">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="06025-235">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-235">- Taskpane</span></span><br><span data-ttu-id="06025-236">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-236">
        - Content</span></span><br><span data-ttu-id="06025-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="06025-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="06025-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="06025-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="06025-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="06025-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="06025-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="06025-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="06025-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="06025-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="06025-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-247">-BindingEvents</span></span><br><span data-ttu-id="06025-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-248">
        -CompressedFile</span></span><br><span data-ttu-id="06025-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-249">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-250">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-250">
        - File</span></span><br><span data-ttu-id="06025-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-251">
        -ImageCoercion</span></span><br><span data-ttu-id="06025-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-252">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-254">
        -PdfFile</span></span><br><span data-ttu-id="06025-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-255">
        - Selection</span></span><br><span data-ttu-id="06025-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-256">
        - Settings</span></span><br><span data-ttu-id="06025-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-257">
        -TableBindings</span></span><br><span data-ttu-id="06025-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-258">
        -TableCoercion</span></span><br><span data-ttu-id="06025-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-259">
        -TextBindings</span></span><br><span data-ttu-id="06025-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-261">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="06025-262">- Область задач</span><span class="sxs-lookup"><span data-stu-id="06025-262">- Taskpane</span></span><br><span data-ttu-id="06025-263">
        - Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-263">
        - Content</span></span><br><span data-ttu-id="06025-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="06025-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="06025-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="06025-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="06025-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="06025-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="06025-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="06025-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="06025-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="06025-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="06025-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="06025-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-274">-BindingEvents</span></span><br><span data-ttu-id="06025-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-275">
        -CompressedFile</span></span><br><span data-ttu-id="06025-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-276">
        -DocumentEvents</span></span><br><span data-ttu-id="06025-277">
        - File</span><span class="sxs-lookup"><span data-stu-id="06025-277">
        - File</span></span><br><span data-ttu-id="06025-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-278">
        -ImageCoercion</span></span><br><span data-ttu-id="06025-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-279">
        -MatrixBindings</span></span><br><span data-ttu-id="06025-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="06025-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-281">
        -PdfFile</span></span><br><span data-ttu-id="06025-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-282">
        - Selection</span></span><br><span data-ttu-id="06025-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-283">
        - Settings</span></span><br><span data-ttu-id="06025-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-284">
        -TableBindings</span></span><br><span data-ttu-id="06025-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-285">
        -TableCoercion</span></span><br><span data-ttu-id="06025-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-286">
        -TextBindings</span></span><br><span data-ttu-id="06025-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="06025-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="06025-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="06025-289">Платформа</span><span class="sxs-lookup"><span data-stu-id="06025-289">Platform</span></span></th>
    <th><span data-ttu-id="06025-290">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="06025-290">Extension points</span></span></th>
    <th><span data-ttu-id="06025-291">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="06025-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="06025-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="06025-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="06025-293">Office Online</span></span></td>
    <td> <span data-ttu-id="06025-294">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-294">- Mail Read</span></span><br><span data-ttu-id="06025-295">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="06025-295">
      - Mail Compose</span></span><br><span data-ttu-id="06025-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="06025-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="06025-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="06025-304">Недоступна</span><span class="sxs-lookup"><span data-stu-id="06025-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-305">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="06025-306">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-306">- Mail Read</span></span><br><span data-ttu-id="06025-307">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="06025-307">
      - Mail Compose</span></span><br><span data-ttu-id="06025-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="06025-313">Недоступно</span><span class="sxs-lookup"><span data-stu-id="06025-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-314">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="06025-315">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-315">- Mail Read</span></span><br><span data-ttu-id="06025-316">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="06025-316">
      - Mail Compose</span></span><br><span data-ttu-id="06025-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="06025-318">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="06025-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="06025-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="06025-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="06025-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="06025-326">Недоступна</span><span class="sxs-lookup"><span data-stu-id="06025-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-327">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="06025-328">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-328">- Mail Read</span></span><br><span data-ttu-id="06025-329">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="06025-329">
      - Mail Compose</span></span><br><span data-ttu-id="06025-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="06025-331">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="06025-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="06025-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="06025-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="06025-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="06025-339">Недоступно</span><span class="sxs-lookup"><span data-stu-id="06025-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-340">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="06025-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="06025-341">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-341">- Mail Read</span></span><br><span data-ttu-id="06025-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="06025-348">Недоступно</span><span class="sxs-lookup"><span data-stu-id="06025-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-349">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="06025-350">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-350">- Mail Read</span></span><br><span data-ttu-id="06025-351">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="06025-351">
      - Mail Compose</span></span><br><span data-ttu-id="06025-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="06025-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="06025-359">Недоступно</span><span class="sxs-lookup"><span data-stu-id="06025-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-360">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="06025-361">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-361">- Mail Read</span></span><br><span data-ttu-id="06025-362">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="06025-362">
      - Mail Compose</span></span><br><span data-ttu-id="06025-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="06025-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="06025-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="06025-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="06025-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="06025-371">Недоступно</span><span class="sxs-lookup"><span data-stu-id="06025-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-372">Office для Android</span><span class="sxs-lookup"><span data-stu-id="06025-372">Office for Android</span></span></td>
    <td> <span data-ttu-id="06025-373">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="06025-373">- Mail Read</span></span><br><span data-ttu-id="06025-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="06025-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="06025-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="06025-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="06025-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="06025-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="06025-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="06025-380">Недоступно</span><span class="sxs-lookup"><span data-stu-id="06025-380">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="06025-381">Word</span><span class="sxs-lookup"><span data-stu-id="06025-381">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="06025-382">Платформа</span><span class="sxs-lookup"><span data-stu-id="06025-382">Platform</span></span></th>
    <th><span data-ttu-id="06025-383">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="06025-383">Extension points</span></span></th>
    <th><span data-ttu-id="06025-384">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="06025-384">API requirement sets</span></span></th>
    <th><span data-ttu-id="06025-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="06025-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="06025-386">Office Online</span><span class="sxs-lookup"><span data-stu-id="06025-386">Office Online</span></span></td>
    <td> <span data-ttu-id="06025-387">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-387">- Taskpane</span></span><br><span data-ttu-id="06025-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="06025-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="06025-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="06025-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-393">-BindingEvents</span></span><br><span data-ttu-id="06025-394">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-394">customXmlParts</span></span><br><span data-ttu-id="06025-395">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-395">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-396">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-396">
         - File</span></span><br><span data-ttu-id="06025-397">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-397">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-398">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-398">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-399">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-399">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-400">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-400">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-401">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-401">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-402">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-402">
         -PdfFile</span></span><br><span data-ttu-id="06025-403">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-403">
         - Selection</span></span><br><span data-ttu-id="06025-404">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-404">
         - Settings</span></span><br><span data-ttu-id="06025-405">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-405">
         -TableBindings</span></span><br><span data-ttu-id="06025-406">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-406">
         -TableCoercion</span></span><br><span data-ttu-id="06025-407">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-407">
         -TextBindings</span></span><br><span data-ttu-id="06025-408">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-408">
         -TextCoercion</span></span><br><span data-ttu-id="06025-409">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-409">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-410">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-410">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="06025-411">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-411">- Taskpane</span></span></td>
    <td> <span data-ttu-id="06025-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-413">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-413">-BindingEvents</span></span><br><span data-ttu-id="06025-414">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-414">
         -CompressedFile</span></span><br><span data-ttu-id="06025-415">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-415">customXmlParts</span></span><br><span data-ttu-id="06025-416">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-416">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-417">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-417">
         - File</span></span><br><span data-ttu-id="06025-418">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-418">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-419">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-419">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-420">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-420">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-421">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-421">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-422">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-422">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-423">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-423">
         -PdfFile</span></span><br><span data-ttu-id="06025-424">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-424">
         - Selection</span></span><br><span data-ttu-id="06025-425">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-425">
         - Settings</span></span><br><span data-ttu-id="06025-426">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-426">
         -TableBindings</span></span><br><span data-ttu-id="06025-427">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-427">
         -TableCoercion</span></span><br><span data-ttu-id="06025-428">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-428">
         -TextBindings</span></span><br><span data-ttu-id="06025-429">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-429">
         -TextCoercion</span></span><br><span data-ttu-id="06025-430">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-430">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-431">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-431">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="06025-432">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-432">- Taskpane</span></span><br><span data-ttu-id="06025-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="06025-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="06025-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="06025-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-438">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-438">-BindingEvents</span></span><br><span data-ttu-id="06025-439">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-439">
         -CompressedFile</span></span><br><span data-ttu-id="06025-440">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-440">customXmlParts</span></span><br><span data-ttu-id="06025-441">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-441">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-442">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-442">
         - File</span></span><br><span data-ttu-id="06025-443">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-443">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-444">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-444">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-445">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-445">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-446">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-446">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-447">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-447">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-448">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-448">
         -PdfFile</span></span><br><span data-ttu-id="06025-449">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-449">
         - Selection</span></span><br><span data-ttu-id="06025-450">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-450">
         - Settings</span></span><br><span data-ttu-id="06025-451">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-451">
         -TableBindings</span></span><br><span data-ttu-id="06025-452">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-452">
         -TableCoercion</span></span><br><span data-ttu-id="06025-453">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-453">
         -TextBindings</span></span><br><span data-ttu-id="06025-454">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-454">
         -TextCoercion</span></span><br><span data-ttu-id="06025-455">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-455">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-456">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-456">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="06025-457">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-457">- Taskpane</span></span><br><span data-ttu-id="06025-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="06025-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="06025-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="06025-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-463">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-463">-BindingEvents</span></span><br><span data-ttu-id="06025-464">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-464">
         -CompressedFile</span></span><br><span data-ttu-id="06025-465">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-465">customXmlParts</span></span><br><span data-ttu-id="06025-466">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-466">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-467">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-467">
         - File</span></span><br><span data-ttu-id="06025-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-469">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-470">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-470">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-471">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-471">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-472">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-472">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-473">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-473">
         -PdfFile</span></span><br><span data-ttu-id="06025-474">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-474">
         - Selection</span></span><br><span data-ttu-id="06025-475">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-475">
         - Settings</span></span><br><span data-ttu-id="06025-476">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-476">
         -TableBindings</span></span><br><span data-ttu-id="06025-477">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-477">
         -TableCoercion</span></span><br><span data-ttu-id="06025-478">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-478">
         -TextBindings</span></span><br><span data-ttu-id="06025-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-479">
         -TextCoercion</span></span><br><span data-ttu-id="06025-480">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-480">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-481">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="06025-481">Office for iOS</span></span></td>
    <td> <span data-ttu-id="06025-482">- Область задач</span><span class="sxs-lookup"><span data-stu-id="06025-482">- Taskpane</span></span></td>
    <td> <span data-ttu-id="06025-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="06025-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="06025-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="06025-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="06025-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="06025-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-487">-BindingEvents</span></span><br><span data-ttu-id="06025-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-488">
         -CompressedFile</span></span><br><span data-ttu-id="06025-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-489">customXmlParts</span></span><br><span data-ttu-id="06025-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-490">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-491">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-491">
         - File</span></span><br><span data-ttu-id="06025-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-492">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-493">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-494">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-495">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-496">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-497">
         -PdfFile</span></span><br><span data-ttu-id="06025-498">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-498">
         - Selection</span></span><br><span data-ttu-id="06025-499">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-499">
         - Settings</span></span><br><span data-ttu-id="06025-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-500">
         -TableBindings</span></span><br><span data-ttu-id="06025-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-501">
         -TableCoercion</span></span><br><span data-ttu-id="06025-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-502">
         -TextBindings</span></span><br><span data-ttu-id="06025-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-503">
         -TextCoercion</span></span><br><span data-ttu-id="06025-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-504">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-505">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-505">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="06025-506">- Область задач</span><span class="sxs-lookup"><span data-stu-id="06025-506">- Taskpane</span></span><br><span data-ttu-id="06025-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="06025-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="06025-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="06025-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="06025-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="06025-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-512">-BindingEvents</span></span><br><span data-ttu-id="06025-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-513">
         -CompressedFile</span></span><br><span data-ttu-id="06025-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-514">customXmlParts</span></span><br><span data-ttu-id="06025-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-515">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-516">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-516">
         - File</span></span><br><span data-ttu-id="06025-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-517">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-518">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-519">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-520">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-521">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-522">
         -PdfFile</span></span><br><span data-ttu-id="06025-523">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-523">
         - Selection</span></span><br><span data-ttu-id="06025-524">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-524">
         - Settings</span></span><br><span data-ttu-id="06025-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-525">
         -TableBindings</span></span><br><span data-ttu-id="06025-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-526">
         -TableCoercion</span></span><br><span data-ttu-id="06025-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-527">
         -TextBindings</span></span><br><span data-ttu-id="06025-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-528">
         -TextCoercion</span></span><br><span data-ttu-id="06025-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-529">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-530">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-530">Office for Mac</span></span></td>
    <td> <span data-ttu-id="06025-531">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-531">- Taskpane</span></span><br><span data-ttu-id="06025-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="06025-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="06025-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="06025-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="06025-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="06025-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="06025-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="06025-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="06025-537">-BindingEvents</span></span><br><span data-ttu-id="06025-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-538">
         -CompressedFile</span></span><br><span data-ttu-id="06025-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="06025-539">customXmlParts</span></span><br><span data-ttu-id="06025-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-540">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-541">
         - File</span></span><br><span data-ttu-id="06025-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-542">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-543">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="06025-544">
         -MatrixBindings</span></span><br><span data-ttu-id="06025-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-545">
         -MatrixCoercion</span></span><br><span data-ttu-id="06025-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-546">
         -OoxmlCoercion</span></span><br><span data-ttu-id="06025-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-547">
         -PdfFile</span></span><br><span data-ttu-id="06025-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-548">
         - Selection</span></span><br><span data-ttu-id="06025-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-549">
         - Settings</span></span><br><span data-ttu-id="06025-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="06025-550">
         -TableBindings</span></span><br><span data-ttu-id="06025-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-551">
         -TableCoercion</span></span><br><span data-ttu-id="06025-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="06025-552">
         -TextBindings</span></span><br><span data-ttu-id="06025-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-553">
         -TextCoercion</span></span><br><span data-ttu-id="06025-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="06025-554">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="06025-555">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="06025-555">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="06025-556">Платформа</span><span class="sxs-lookup"><span data-stu-id="06025-556">Platform</span></span></th>
    <th><span data-ttu-id="06025-557">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="06025-557">Extension points</span></span></th>
    <th><span data-ttu-id="06025-558">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="06025-558">API requirement sets</span></span></th>
    <th><span data-ttu-id="06025-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="06025-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="06025-560">Office Online</span><span class="sxs-lookup"><span data-stu-id="06025-560">Office Online</span></span></td>
    <td> <span data-ttu-id="06025-561">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-561">- Content</span></span><br><span data-ttu-id="06025-562">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-562">
         - Taskpane</span></span><br><span data-ttu-id="06025-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-565">-ActiveView</span></span><br><span data-ttu-id="06025-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-566">
         -CompressedFile</span></span><br><span data-ttu-id="06025-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-567">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-568">
         - File</span></span><br><span data-ttu-id="06025-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-569">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-570">
         -PdfFile</span></span><br><span data-ttu-id="06025-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-571">
         - Selection</span></span><br><span data-ttu-id="06025-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-572">
         - Settings</span></span><br><span data-ttu-id="06025-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-573">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-574">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-574">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="06025-575">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-575">- Content</span></span><br><span data-ttu-id="06025-576">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-576">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="06025-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="06025-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="06025-578">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-578">-ActiveView</span></span><br><span data-ttu-id="06025-579">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-579">
         -CompressedFile</span></span><br><span data-ttu-id="06025-580">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-580">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-581">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-581">
         - File</span></span><br><span data-ttu-id="06025-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-582">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-583">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-583">
         -PdfFile</span></span><br><span data-ttu-id="06025-584">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-584">
         - Selection</span></span><br><span data-ttu-id="06025-585">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-585">
         - Settings</span></span><br><span data-ttu-id="06025-586">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-586">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-587">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-587">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="06025-588">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-588">- Content</span></span><br><span data-ttu-id="06025-589">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-589">
         - Taskpane</span></span><br><span data-ttu-id="06025-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-592">-ActiveView</span></span><br><span data-ttu-id="06025-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-593">
         -CompressedFile</span></span><br><span data-ttu-id="06025-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-594">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-595">
         - File</span></span><br><span data-ttu-id="06025-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-596">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-597">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-597">
         -PdfFile</span></span><br><span data-ttu-id="06025-598">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-598">
         - Selection</span></span><br><span data-ttu-id="06025-599">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-599">
         - Settings</span></span><br><span data-ttu-id="06025-600">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-600">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-601">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="06025-601">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="06025-602">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-602">- Content</span></span><br><span data-ttu-id="06025-603">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-603">
         - Taskpane</span></span><br><span data-ttu-id="06025-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-606">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-606">-ActiveView</span></span><br><span data-ttu-id="06025-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-607">
         -CompressedFile</span></span><br><span data-ttu-id="06025-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-608">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-609">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-609">
         - File</span></span><br><span data-ttu-id="06025-610">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-610">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-611">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-611">
         -PdfFile</span></span><br><span data-ttu-id="06025-612">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-612">
         - Selection</span></span><br><span data-ttu-id="06025-613">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-613">
         - Settings</span></span><br><span data-ttu-id="06025-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-614">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-615">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="06025-615">Office for iOS</span></span></td>
    <td> <span data-ttu-id="06025-616">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-616">- Content</span></span><br><span data-ttu-id="06025-617">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-617">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="06025-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="06025-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-619">-ActiveView</span></span><br><span data-ttu-id="06025-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-620">
         -CompressedFile</span></span><br><span data-ttu-id="06025-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-621">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-622">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-622">
         - File</span></span><br><span data-ttu-id="06025-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-623">
         -PdfFile</span></span><br><span data-ttu-id="06025-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-624">
         - Selection</span></span><br><span data-ttu-id="06025-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-625">
         - Settings</span></span><br><span data-ttu-id="06025-626">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-626">
         -TextCoercion</span></span><br><span data-ttu-id="06025-627">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-627">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-628">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-628">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="06025-629">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-629">- Content</span></span><br><span data-ttu-id="06025-630">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-630">
         - Taskpane</span></span><br><span data-ttu-id="06025-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-633">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-633">-ActiveView</span></span><br><span data-ttu-id="06025-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-634">
         -CompressedFile</span></span><br><span data-ttu-id="06025-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-635">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-636">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-636">
         - File</span></span><br><span data-ttu-id="06025-637">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-637">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-638">
         -PdfFile</span></span><br><span data-ttu-id="06025-639">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-639">
         - Selection</span></span><br><span data-ttu-id="06025-640">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-640">
         - Settings</span></span><br><span data-ttu-id="06025-641">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-641">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="06025-642">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="06025-642">Office for Mac</span></span></td>
    <td> <span data-ttu-id="06025-643">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-643">- Content</span></span><br><span data-ttu-id="06025-644">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-644">
         - Taskpane</span></span><br><span data-ttu-id="06025-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-647">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="06025-647">-ActiveView</span></span><br><span data-ttu-id="06025-648">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="06025-648">
         -CompressedFile</span></span><br><span data-ttu-id="06025-649">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-649">
         -DocumentEvents</span></span><br><span data-ttu-id="06025-650">
         - File</span><span class="sxs-lookup"><span data-stu-id="06025-650">
         - File</span></span><br><span data-ttu-id="06025-651">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-651">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-652">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="06025-652">
         -PdfFile</span></span><br><span data-ttu-id="06025-653">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="06025-653">
         - Selection</span></span><br><span data-ttu-id="06025-654">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-654">
         - Settings</span></span><br><span data-ttu-id="06025-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-655">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="06025-656">OneNote</span><span class="sxs-lookup"><span data-stu-id="06025-656">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="06025-657">Платформа</span><span class="sxs-lookup"><span data-stu-id="06025-657">Platform</span></span></th>
    <th><span data-ttu-id="06025-658">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="06025-658">Extension points</span></span></th>
    <th><span data-ttu-id="06025-659">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="06025-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="06025-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="06025-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="06025-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="06025-661">Office Online</span></span></td>
    <td> <span data-ttu-id="06025-662">- Контентные</span><span class="sxs-lookup"><span data-stu-id="06025-662">- Content</span></span><br><span data-ttu-id="06025-663">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="06025-663">
         - Taskpane</span></span><br><span data-ttu-id="06025-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="06025-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="06025-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="06025-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="06025-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="06025-667">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="06025-667">-DocumentEvents</span></span><br><span data-ttu-id="06025-668">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-668">
         -HtmlCoercion</span></span><br><span data-ttu-id="06025-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-669">
         -ImageCoercion</span></span><br><span data-ttu-id="06025-670">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="06025-670">
         - Settings</span></span><br><span data-ttu-id="06025-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="06025-671">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="06025-672">См. также</span><span class="sxs-lookup"><span data-stu-id="06025-672">See also</span></span>

- [<span data-ttu-id="06025-673">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="06025-673">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="06025-674">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="06025-674">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="06025-675">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="06025-675">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="06025-676">Ссылка на API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="06025-676">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
