---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов  для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: bc7ac5c97c041a546c160c05cffc2c80db1ff1b1
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506352"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e15cf-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e15cf-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e15cf-p101">Работа надстройки Office должным образом может зависеть от ведущего приложения Office, набора требований, элемента или версии API. В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API,  которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="e15cf-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="e15cf-p102">Если ячейка таблицы содержит символ звездочки (\*), это означает, что поддержка скоро появится. С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e15cf-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="e15cf-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="e15cf-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e15cf-110">Excel</span><span class="sxs-lookup"><span data-stu-id="e15cf-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e15cf-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="e15cf-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e15cf-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="e15cf-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e15cf-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e15cf-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e15cf-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="e15cf-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="e15cf-115">Office Online</span></span></td>
    <td> <span data-ttu-id="e15cf-116">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-116">- Taskpane</span></span><br><span data-ttu-id="e15cf-117">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-117">
        - Content</span></span><br><span data-ttu-id="e15cf-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a>
    </span><span class="sxs-lookup"><span data-stu-id="e15cf-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e15cf-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e15cf-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e15cf-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e15cf-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e15cf-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e15cf-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e15cf-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e15cf-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e15cf-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-128">
        -BindingEvents</span></span><br><span data-ttu-id="e15cf-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-129">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-130">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-131">
        - File</span></span><br><span data-ttu-id="e15cf-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-132">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-134">
        - Selection</span></span><br><span data-ttu-id="e15cf-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-135">
        - Settings</span></span><br><span data-ttu-id="e15cf-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-136">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-137">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-138">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-140">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e15cf-141">
        - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-141">
        - Taskpane</span></span><br><span data-ttu-id="e15cf-142">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e15cf-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-144">
        -BindingEvents</span></span><br><span data-ttu-id="e15cf-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-145">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-146">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-147">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-147">
        - File</span></span><br><span data-ttu-id="e15cf-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-148">
        -ImageCoercion</span></span><br><span data-ttu-id="e15cf-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-149">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-151">
        - Selection</span></span><br><span data-ttu-id="e15cf-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-152">
        - Settings</span></span><br><span data-ttu-id="e15cf-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-153">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-154">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-155">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-157">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e15cf-158">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-158">- Taskpane</span></span><br><span data-ttu-id="e15cf-159">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-159">
        - Content</span></span><br><span data-ttu-id="e15cf-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e15cf-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e15cf-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e15cf-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e15cf-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e15cf-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e15cf-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e15cf-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e15cf-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e15cf-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-170">-BindingEvents</span></span><br><span data-ttu-id="e15cf-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-171">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-172">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-173">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-173">
        - File</span></span><br><span data-ttu-id="e15cf-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-174">
        -ImageCoercion</span></span><br><span data-ttu-id="e15cf-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-175">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-177">
        - Selection</span></span><br><span data-ttu-id="e15cf-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-178">
        - Settings</span></span><br><span data-ttu-id="e15cf-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-179">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-180">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-181">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-183">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="e15cf-184">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-184">- Taskpane</span></span><br><span data-ttu-id="e15cf-185">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-185">
        - Content</span></span><br><span data-ttu-id="e15cf-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e15cf-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e15cf-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e15cf-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e15cf-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e15cf-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e15cf-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e15cf-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e15cf-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e15cf-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-196">-BindingEvents</span></span><br><span data-ttu-id="e15cf-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-197">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-198">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-199">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-199">
        - File</span></span><br><span data-ttu-id="e15cf-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-200">
        -ImageCoercion</span></span><br><span data-ttu-id="e15cf-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-201">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-203">
        - Selection</span></span><br><span data-ttu-id="e15cf-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-204">
        - Settings</span></span><br><span data-ttu-id="e15cf-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-205">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-206">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-207">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-209">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="e15cf-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="e15cf-210">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-210">- Taskpane</span></span><br><span data-ttu-id="e15cf-211">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-211">
        - Content</span></span></td>
    <td><span data-ttu-id="e15cf-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e15cf-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e15cf-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e15cf-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e15cf-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e15cf-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e15cf-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e15cf-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e15cf-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-221">-BindingEvents</span></span><br><span data-ttu-id="e15cf-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-222">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-223">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-224">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-224">
        - File</span></span><br><span data-ttu-id="e15cf-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-225">
        -ImageCoercion</span></span><br><span data-ttu-id="e15cf-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-226">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-228">
        - Selection</span></span><br><span data-ttu-id="e15cf-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-229">
        - Settings</span></span><br><span data-ttu-id="e15cf-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-230">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-231">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-232">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-234">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e15cf-235">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-235">- Taskpane</span></span><br><span data-ttu-id="e15cf-236">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-236">
        - Content</span></span><br><span data-ttu-id="e15cf-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e15cf-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e15cf-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e15cf-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e15cf-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e15cf-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e15cf-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e15cf-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e15cf-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e15cf-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-247">-BindingEvents</span></span><br><span data-ttu-id="e15cf-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-248">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-249">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-250">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-250">
        - File</span></span><br><span data-ttu-id="e15cf-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-251">
        -ImageCoercion</span></span><br><span data-ttu-id="e15cf-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-252">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-254">
        -PdfFile</span></span><br><span data-ttu-id="e15cf-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-255">
        - Selection</span></span><br><span data-ttu-id="e15cf-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-256">
        - Settings</span></span><br><span data-ttu-id="e15cf-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-257">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-258">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-259">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-261">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="e15cf-262">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-262">- Taskpane</span></span><br><span data-ttu-id="e15cf-263">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-263">
        - Content</span></span><br><span data-ttu-id="e15cf-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e15cf-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e15cf-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e15cf-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e15cf-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e15cf-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e15cf-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e15cf-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e15cf-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e15cf-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e15cf-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-274">-BindingEvents</span></span><br><span data-ttu-id="e15cf-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-275">
        -CompressedFile</span></span><br><span data-ttu-id="e15cf-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-276">
        -DocumentEvents</span></span><br><span data-ttu-id="e15cf-277">
        - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-277">
        - File</span></span><br><span data-ttu-id="e15cf-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-278">
        -ImageCoercion</span></span><br><span data-ttu-id="e15cf-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-279">
        -MatrixBindings</span></span><br><span data-ttu-id="e15cf-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-281">
        -PdfFile</span></span><br><span data-ttu-id="e15cf-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-282">
        - Selection</span></span><br><span data-ttu-id="e15cf-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-283">
        - Settings</span></span><br><span data-ttu-id="e15cf-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-284">
        -TableBindings</span></span><br><span data-ttu-id="e15cf-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-285">
        -TableCoercion</span></span><br><span data-ttu-id="e15cf-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-286">
        -TextBindings</span></span><br><span data-ttu-id="e15cf-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e15cf-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="e15cf-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e15cf-289">Платформа</span><span class="sxs-lookup"><span data-stu-id="e15cf-289">Platform</span></span></th>
    <th><span data-ttu-id="e15cf-290">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="e15cf-290">Extension points</span></span></th>
    <th><span data-ttu-id="e15cf-291">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e15cf-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="e15cf-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="e15cf-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="e15cf-293">Office Online</span></span></td>
    <td> <span data-ttu-id="e15cf-294">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-294">- Mail Read</span></span><br><span data-ttu-id="e15cf-295">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-295">
      - Mail Compose</span></span><br><span data-ttu-id="e15cf-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e15cf-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e15cf-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e15cf-304">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-305">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e15cf-306">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-306">- Mail Read</span></span><br><span data-ttu-id="e15cf-307">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-307">
      - Mail Compose</span></span><br><span data-ttu-id="e15cf-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e15cf-313">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-314">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e15cf-315">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-315">- Mail Read</span></span><br><span data-ttu-id="e15cf-316">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-316">
      - Mail Compose</span></span><br><span data-ttu-id="e15cf-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e15cf-318">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="e15cf-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e15cf-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e15cf-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e15cf-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e15cf-326">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-327">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="e15cf-328">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-328">- Mail Read</span></span><br><span data-ttu-id="e15cf-329">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-329">
      - Mail Compose</span></span><br><span data-ttu-id="e15cf-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e15cf-331">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="e15cf-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e15cf-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e15cf-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e15cf-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e15cf-339">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-340">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="e15cf-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e15cf-341">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-341">- Mail Read</span></span><br><span data-ttu-id="e15cf-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e15cf-348">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-349">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e15cf-350">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-350">- Mail Read</span></span><br><span data-ttu-id="e15cf-351">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-351">
      - Mail Compose</span></span><br><span data-ttu-id="e15cf-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e15cf-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e15cf-359">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-360">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="e15cf-361">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-361">- Mail Read</span></span><br><span data-ttu-id="e15cf-362">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-362">
      - Mail Compose</span></span><br><span data-ttu-id="e15cf-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e15cf-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e15cf-370">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-371">Office для Android</span><span class="sxs-lookup"><span data-stu-id="e15cf-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="e15cf-372">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="e15cf-372">- Mail Read</span></span><br><span data-ttu-id="e15cf-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e15cf-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e15cf-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e15cf-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e15cf-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e15cf-379">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e15cf-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e15cf-380">Word</span><span class="sxs-lookup"><span data-stu-id="e15cf-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e15cf-381">Платформа</span><span class="sxs-lookup"><span data-stu-id="e15cf-381">Platform</span></span></th>
    <th><span data-ttu-id="e15cf-382">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="e15cf-382">Extension points</span></span></th>
    <th><span data-ttu-id="e15cf-383">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e15cf-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="e15cf-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="e15cf-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="e15cf-385">Office Online</span></span></td>
    <td> <span data-ttu-id="e15cf-386">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-386">- Taskpane</span></span><br><span data-ttu-id="e15cf-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e15cf-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e15cf-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e15cf-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-392">-BindingEvents</span></span><br><span data-ttu-id="e15cf-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-393">customXmlParts</span></span><br><span data-ttu-id="e15cf-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-394">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-395">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-395">
         - File</span></span><br><span data-ttu-id="e15cf-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-397">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-398">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-401">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-402">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-402">
         - Selection</span></span><br><span data-ttu-id="e15cf-403">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-403">
         - Settings</span></span><br><span data-ttu-id="e15cf-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-404">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-405">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-406">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-407">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-409">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e15cf-410">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e15cf-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-412">-BindingEvents</span></span><br><span data-ttu-id="e15cf-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-413">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-414">customXmlParts</span></span><br><span data-ttu-id="e15cf-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-415">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-416">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-416">
         - File</span></span><br><span data-ttu-id="e15cf-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-418">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-419">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-422">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-423">
         - Selection</span></span><br><span data-ttu-id="e15cf-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-424">
         - Settings</span></span><br><span data-ttu-id="e15cf-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-425">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-426">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-427">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-428">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-430">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e15cf-431">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-431">- Taskpane</span></span><br><span data-ttu-id="e15cf-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e15cf-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e15cf-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e15cf-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-437">-BindingEvents</span></span><br><span data-ttu-id="e15cf-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-438">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-439">customXmlParts</span></span><br><span data-ttu-id="e15cf-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-440">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-441">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-441">
         - File</span></span><br><span data-ttu-id="e15cf-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-443">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-444">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-447">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-448">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-448">
         - Selection</span></span><br><span data-ttu-id="e15cf-449">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-449">
         - Settings</span></span><br><span data-ttu-id="e15cf-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-450">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-451">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-452">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-453">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-455">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="e15cf-456">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-456">- Taskpane</span></span><br><span data-ttu-id="e15cf-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e15cf-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e15cf-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e15cf-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-462">-BindingEvents</span></span><br><span data-ttu-id="e15cf-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-463">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-464">customXmlParts</span></span><br><span data-ttu-id="e15cf-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-465">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-466">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-466">
         - File</span></span><br><span data-ttu-id="e15cf-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-468">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-469">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-472">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-473">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-473">
         - Selection</span></span><br><span data-ttu-id="e15cf-474">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-474">
         - Settings</span></span><br><span data-ttu-id="e15cf-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-475">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-476">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-477">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-478">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-480">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="e15cf-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e15cf-481">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e15cf-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e15cf-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e15cf-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e15cf-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e15cf-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e15cf-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-486">-BindingEvents</span></span><br><span data-ttu-id="e15cf-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-487">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-488">customXmlParts</span></span><br><span data-ttu-id="e15cf-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-489">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-490">
         - File</span></span><br><span data-ttu-id="e15cf-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-492">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-493">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-496">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-497">
         - Selection</span></span><br><span data-ttu-id="e15cf-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-498">
         - Settings</span></span><br><span data-ttu-id="e15cf-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-499">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-500">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-501">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-502">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-504">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e15cf-505">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-505">- Taskpane</span></span><br><span data-ttu-id="e15cf-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e15cf-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e15cf-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e15cf-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e15cf-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e15cf-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-511">-BindingEvents</span></span><br><span data-ttu-id="e15cf-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-512">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-513">customXmlParts</span></span><br><span data-ttu-id="e15cf-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-514">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-515">
         - File</span></span><br><span data-ttu-id="e15cf-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-517">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-518">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-521">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-522">
         - Selection</span></span><br><span data-ttu-id="e15cf-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-523">
         - Settings</span></span><br><span data-ttu-id="e15cf-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-524">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-525">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-526">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-527">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-529">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="e15cf-530">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-530">- Taskpane</span></span><br><span data-ttu-id="e15cf-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e15cf-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e15cf-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e15cf-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e15cf-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e15cf-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-536">-BindingEvents</span></span><br><span data-ttu-id="e15cf-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-537">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e15cf-538">customXmlParts</span></span><br><span data-ttu-id="e15cf-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-539">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-540">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-540">
         - File</span></span><br><span data-ttu-id="e15cf-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-542">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-543">
         -MatrixBindings</span></span><br><span data-ttu-id="e15cf-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="e15cf-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e15cf-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-546">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-547">
         - Selection</span></span><br><span data-ttu-id="e15cf-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-548">
         - Settings</span></span><br><span data-ttu-id="e15cf-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-549">
         -TableBindings</span></span><br><span data-ttu-id="e15cf-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-550">
         -TableCoercion</span></span><br><span data-ttu-id="e15cf-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e15cf-551">
         -TextBindings</span></span><br><span data-ttu-id="e15cf-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-552">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e15cf-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e15cf-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e15cf-555">Платформа</span><span class="sxs-lookup"><span data-stu-id="e15cf-555">Platform</span></span></th>
    <th><span data-ttu-id="e15cf-556">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="e15cf-556">Extension points</span></span></th>
    <th><span data-ttu-id="e15cf-557">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e15cf-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="e15cf-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="e15cf-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="e15cf-559">Office Online</span></span></td>
    <td> <span data-ttu-id="e15cf-560">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-560">- Content</span></span><br><span data-ttu-id="e15cf-561">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-561">
         - Taskpane</span></span><br><span data-ttu-id="e15cf-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-564">-ActiveView</span></span><br><span data-ttu-id="e15cf-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-565">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-566">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-567">
         - File</span></span><br><span data-ttu-id="e15cf-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-568">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-569">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-570">
         - Selection</span></span><br><span data-ttu-id="e15cf-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-571">
         - Settings</span></span><br><span data-ttu-id="e15cf-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-573">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e15cf-574">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-574">- Content</span></span><br><span data-ttu-id="e15cf-575">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="e15cf-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e15cf-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e15cf-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-577">-ActiveView</span></span><br><span data-ttu-id="e15cf-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-578">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-579">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-580">
         - File</span></span><br><span data-ttu-id="e15cf-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-581">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-582">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-583">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-583">
         - Selection</span></span><br><span data-ttu-id="e15cf-584">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-584">
         - Settings</span></span><br><span data-ttu-id="e15cf-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-586">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e15cf-587">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-587">- Content</span></span><br><span data-ttu-id="e15cf-588">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-588">
         - Taskpane</span></span><br><span data-ttu-id="e15cf-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-591">-ActiveView</span></span><br><span data-ttu-id="e15cf-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-592">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-593">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-594">
         - File</span></span><br><span data-ttu-id="e15cf-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-595">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-596">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-597">
         - Selection</span></span><br><span data-ttu-id="e15cf-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-598">
         - Settings</span></span><br><span data-ttu-id="e15cf-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-600">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="e15cf-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="e15cf-601">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-601">- Content</span></span><br><span data-ttu-id="e15cf-602">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-602">
         - Taskpane</span></span><br><span data-ttu-id="e15cf-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-605">-ActiveView</span></span><br><span data-ttu-id="e15cf-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-606">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-607">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-608">
         - File</span></span><br><span data-ttu-id="e15cf-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-609">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-610">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-611">
         - Selection</span></span><br><span data-ttu-id="e15cf-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-612">
         - Settings</span></span><br><span data-ttu-id="e15cf-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-614">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="e15cf-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e15cf-615">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-615">- Content</span></span><br><span data-ttu-id="e15cf-616">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="e15cf-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e15cf-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-618">-ActiveView</span></span><br><span data-ttu-id="e15cf-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-619">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-620">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-621">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-621">
         - File</span></span><br><span data-ttu-id="e15cf-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-622">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-623">
         - Selection</span></span><br><span data-ttu-id="e15cf-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-624">
         - Settings</span></span><br><span data-ttu-id="e15cf-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-625">
         -TextCoercion</span></span><br><span data-ttu-id="e15cf-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-627">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e15cf-628">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-628">- Content</span></span><br><span data-ttu-id="e15cf-629">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-629">
         - Taskpane</span></span><br><span data-ttu-id="e15cf-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-632">-ActiveView</span></span><br><span data-ttu-id="e15cf-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-633">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-634">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-635">
         - File</span></span><br><span data-ttu-id="e15cf-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-636">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-637">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-638">
         - Selection</span></span><br><span data-ttu-id="e15cf-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-639">
         - Settings</span></span><br><span data-ttu-id="e15cf-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-641">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="e15cf-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="e15cf-642">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-642">- Content</span></span><br><span data-ttu-id="e15cf-643">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-643">
         - Taskpane</span></span><br><span data-ttu-id="e15cf-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e15cf-646">-ActiveView</span></span><br><span data-ttu-id="e15cf-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-647">
         -CompressedFile</span></span><br><span data-ttu-id="e15cf-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-648">
         -DocumentEvents</span></span><br><span data-ttu-id="e15cf-649">
         - File</span><span class="sxs-lookup"><span data-stu-id="e15cf-649">
         - File</span></span><br><span data-ttu-id="e15cf-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-650">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e15cf-651">
         -PdfFile</span></span><br><span data-ttu-id="e15cf-652">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e15cf-652">
         - Selection</span></span><br><span data-ttu-id="e15cf-653">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-653">
         - Settings</span></span><br><span data-ttu-id="e15cf-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e15cf-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="e15cf-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e15cf-656">Платформа</span><span class="sxs-lookup"><span data-stu-id="e15cf-656">Platform</span></span></th>
    <th><span data-ttu-id="e15cf-657">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="e15cf-657">Extension points</span></span></th>
    <th><span data-ttu-id="e15cf-658">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e15cf-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="e15cf-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="e15cf-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e15cf-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="e15cf-660">Office Online</span></span></td>
    <td> <span data-ttu-id="e15cf-661">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="e15cf-661">- Content</span></span><br><span data-ttu-id="e15cf-662">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="e15cf-662">
         - Taskpane</span></span><br><span data-ttu-id="e15cf-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e15cf-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e15cf-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e15cf-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e15cf-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e15cf-666">-DocumentEvents</span></span><br><span data-ttu-id="e15cf-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="e15cf-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-668">
         -ImageCoercion</span></span><br><span data-ttu-id="e15cf-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e15cf-669">
         - Settings</span></span><br><span data-ttu-id="e15cf-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e15cf-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e15cf-671">См. также</span><span class="sxs-lookup"><span data-stu-id="e15cf-671">See also</span></span>

- [<span data-ttu-id="e15cf-672">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e15cf-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e15cf-673">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="e15cf-673">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e15cf-674">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="e15cf-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e15cf-675">Ссылка на API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="e15cf-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
