---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов  для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 6f7b5b565773457e6cd8a9eee69eb304784a29a9
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459317"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b6e77-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b6e77-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b6e77-p101">Работа надстройки Office должным образом может зависеть от ведущего приложения Office, набора требований, элемента или версии API. В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API,  которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="b6e77-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="b6e77-p102">Если ячейка таблицы содержит символ звездочки (\*), это означает, что поддержка скоро появится. С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b6e77-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="b6e77-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="b6e77-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="b6e77-110">Excel</span><span class="sxs-lookup"><span data-stu-id="b6e77-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b6e77-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="b6e77-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b6e77-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b6e77-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b6e77-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b6e77-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b6e77-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e77-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e77-115">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e77-116">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-116">- Taskpane</span></span><br><span data-ttu-id="b6e77-117">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-117">
        - Content</span></span><br><span data-ttu-id="b6e77-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a>
    </span><span class="sxs-lookup"><span data-stu-id="b6e77-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b6e77-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e77-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e77-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e77-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e77-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e77-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e77-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b6e77-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e77-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-128">
        -BindingEvents</span></span><br><span data-ttu-id="b6e77-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-129">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-130">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-131">
        - File</span></span><br><span data-ttu-id="b6e77-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-132">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-134">
        - Selection</span></span><br><span data-ttu-id="b6e77-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-135">
        - Settings</span></span><br><span data-ttu-id="b6e77-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-136">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-137">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-138">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-140">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="b6e77-141">
        - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-141">
        - Taskpane</span></span><br><span data-ttu-id="b6e77-142">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b6e77-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-144">
        -BindingEvents</span></span><br><span data-ttu-id="b6e77-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-145">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-146">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-147">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-147">
        - File</span></span><br><span data-ttu-id="b6e77-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-148">
        -ImageCoercion</span></span><br><span data-ttu-id="b6e77-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-149">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-151">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-151">
        - Selection</span></span><br><span data-ttu-id="b6e77-152">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-152">
        - Settings</span></span><br><span data-ttu-id="b6e77-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-153">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-154">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-155">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-157">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="b6e77-158">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-158">- Taskpane</span></span><br><span data-ttu-id="b6e77-159">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-159">
        - Content</span></span><br><span data-ttu-id="b6e77-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e77-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e77-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e77-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e77-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e77-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e77-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e77-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b6e77-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e77-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-170">-BindingEvents</span></span><br><span data-ttu-id="b6e77-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-171">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-172">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-173">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-173">
        - File</span></span><br><span data-ttu-id="b6e77-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-174">
        -ImageCoercion</span></span><br><span data-ttu-id="b6e77-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-175">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-177">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-177">
        - Selection</span></span><br><span data-ttu-id="b6e77-178">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-178">
        - Settings</span></span><br><span data-ttu-id="b6e77-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-179">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-180">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-181">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-183">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="b6e77-184">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-184">- Taskpane</span></span><br><span data-ttu-id="b6e77-185">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-185">
        - Content</span></span><br><span data-ttu-id="b6e77-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e77-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e77-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e77-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e77-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e77-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e77-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e77-193">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b6e77-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e77-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-196">-BindingEvents</span></span><br><span data-ttu-id="b6e77-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-197">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-198">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-199">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-199">
        - File</span></span><br><span data-ttu-id="b6e77-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-200">
        -ImageCoercion</span></span><br><span data-ttu-id="b6e77-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-201">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-203">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-203">
        - Selection</span></span><br><span data-ttu-id="b6e77-204">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-204">
        - Settings</span></span><br><span data-ttu-id="b6e77-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-205">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-206">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-207">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-209">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="b6e77-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="b6e77-210">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-210">- Taskpane</span></span><br><span data-ttu-id="b6e77-211">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-211">
        - Content</span></span></td>
    <td><span data-ttu-id="b6e77-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e77-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e77-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e77-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e77-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e77-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e77-218">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b6e77-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e77-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-221">-BindingEvents</span></span><br><span data-ttu-id="b6e77-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-222">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-223">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-224">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-224">
        - File</span></span><br><span data-ttu-id="b6e77-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-225">
        -ImageCoercion</span></span><br><span data-ttu-id="b6e77-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-226">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-228">
        - Selection</span></span><br><span data-ttu-id="b6e77-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-229">
        - Settings</span></span><br><span data-ttu-id="b6e77-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-230">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-231">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-232">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-234">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="b6e77-235">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-235">- Taskpane</span></span><br><span data-ttu-id="b6e77-236">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-236">
        - Content</span></span><br><span data-ttu-id="b6e77-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e77-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e77-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e77-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e77-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e77-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e77-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e77-244">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b6e77-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e77-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-247">-BindingEvents</span></span><br><span data-ttu-id="b6e77-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-248">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-249">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-250">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-250">
        - File</span></span><br><span data-ttu-id="b6e77-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-251">
        -ImageCoercion</span></span><br><span data-ttu-id="b6e77-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-252">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-254">
        -PdfFile</span></span><br><span data-ttu-id="b6e77-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-255">
        - Selection</span></span><br><span data-ttu-id="b6e77-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-256">
        - Settings</span></span><br><span data-ttu-id="b6e77-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-257">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-258">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-259">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-261">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="b6e77-262">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-262">- Taskpane</span></span><br><span data-ttu-id="b6e77-263">
        - Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-263">
        - Content</span></span><br><span data-ttu-id="b6e77-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b6e77-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b6e77-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b6e77-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b6e77-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b6e77-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b6e77-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b6e77-271">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b6e77-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b6e77-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b6e77-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-274">-BindingEvents</span></span><br><span data-ttu-id="b6e77-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-275">
        -CompressedFile</span></span><br><span data-ttu-id="b6e77-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-276">
        -DocumentEvents</span></span><br><span data-ttu-id="b6e77-277">
        - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-277">
        - File</span></span><br><span data-ttu-id="b6e77-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-278">
        -ImageCoercion</span></span><br><span data-ttu-id="b6e77-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-279">
        -MatrixBindings</span></span><br><span data-ttu-id="b6e77-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-281">
        -PdfFile</span></span><br><span data-ttu-id="b6e77-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-282">
        - Selection</span></span><br><span data-ttu-id="b6e77-283">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-283">
        - Settings</span></span><br><span data-ttu-id="b6e77-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-284">
        -TableBindings</span></span><br><span data-ttu-id="b6e77-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-285">
        -TableCoercion</span></span><br><span data-ttu-id="b6e77-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-286">
        -TextBindings</span></span><br><span data-ttu-id="b6e77-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="b6e77-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="b6e77-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e77-289">Платформа</span><span class="sxs-lookup"><span data-stu-id="b6e77-289">Platform</span></span></th>
    <th><span data-ttu-id="b6e77-290">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b6e77-290">Extension points</span></span></th>
    <th><span data-ttu-id="b6e77-291">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b6e77-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e77-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e77-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e77-293">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e77-294">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-294">- Mail Read</span></span><br><span data-ttu-id="b6e77-295">
      -  Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-295">
      - Mail Compose</span></span><br><span data-ttu-id="b6e77-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e77-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b6e77-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b6e77-304">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-305">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e77-306">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-306">- Mail Read</span></span><br><span data-ttu-id="b6e77-307">
      -  Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-307">
      - Mail Compose</span></span><br><span data-ttu-id="b6e77-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b6e77-313">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-314">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e77-315">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-315">- Mail Read</span></span><br><span data-ttu-id="b6e77-316">
      -  Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-316">
      - Mail Compose</span></span><br><span data-ttu-id="b6e77-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b6e77-318">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="b6e77-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b6e77-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e77-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b6e77-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b6e77-326">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-327">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="b6e77-328">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-328">- Mail Read</span></span><br><span data-ttu-id="b6e77-329">
      -  Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-329">
      - Mail Compose</span></span><br><span data-ttu-id="b6e77-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b6e77-331">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="b6e77-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b6e77-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e77-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b6e77-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b6e77-339">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-340">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="b6e77-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b6e77-341">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-341">- Mail Read</span></span><br><span data-ttu-id="b6e77-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b6e77-348">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-349">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e77-350">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-350">- Mail Read</span></span><br><span data-ttu-id="b6e77-351">
      -  Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-351">
      - Mail Compose</span></span><br><span data-ttu-id="b6e77-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e77-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b6e77-359">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-360">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="b6e77-361">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-361">- Mail Read</span></span><br><span data-ttu-id="b6e77-362">
      -  Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-362">
      - Mail Compose</span></span><br><span data-ttu-id="b6e77-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b6e77-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b6e77-370">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-371">Office для Android</span><span class="sxs-lookup"><span data-stu-id="b6e77-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="b6e77-372">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="b6e77-372">- Mail Read</span></span><br><span data-ttu-id="b6e77-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b6e77-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b6e77-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b6e77-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b6e77-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b6e77-379">Недоступна</span><span class="sxs-lookup"><span data-stu-id="b6e77-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="b6e77-380">Word</span><span class="sxs-lookup"><span data-stu-id="b6e77-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e77-381">Платформа</span><span class="sxs-lookup"><span data-stu-id="b6e77-381">Platform</span></span></th>
    <th><span data-ttu-id="b6e77-382">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b6e77-382">Extension points</span></span></th>
    <th><span data-ttu-id="b6e77-383">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b6e77-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e77-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e77-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e77-385">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e77-386">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-386">- Taskpane</span></span><br><span data-ttu-id="b6e77-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e77-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e77-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e77-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-392">-BindingEvents</span></span><br><span data-ttu-id="b6e77-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-393">customXmlParts</span></span><br><span data-ttu-id="b6e77-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-394">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-395">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-395">
         - File</span></span><br><span data-ttu-id="b6e77-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-397">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-398">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-401">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-402">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-402">
         - Selection</span></span><br><span data-ttu-id="b6e77-403">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-403">
         - Settings</span></span><br><span data-ttu-id="b6e77-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-404">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-405">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-406">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-407">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-409">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e77-410">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="b6e77-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-412">-BindingEvents</span></span><br><span data-ttu-id="b6e77-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-413">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-414">customXmlParts</span></span><br><span data-ttu-id="b6e77-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-415">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-416">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-416">
         - File</span></span><br><span data-ttu-id="b6e77-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-418">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-419">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-422">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-423">
         - Selection</span></span><br><span data-ttu-id="b6e77-424">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-424">
         - Settings</span></span><br><span data-ttu-id="b6e77-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-425">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-426">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-427">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-428">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-430">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e77-431">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-431">- Taskpane</span></span><br><span data-ttu-id="b6e77-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e77-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e77-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e77-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-437">-BindingEvents</span></span><br><span data-ttu-id="b6e77-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-438">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-439">customXmlParts</span></span><br><span data-ttu-id="b6e77-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-440">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-441">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-441">
         - File</span></span><br><span data-ttu-id="b6e77-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-443">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-444">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-447">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-448">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-448">
         - Selection</span></span><br><span data-ttu-id="b6e77-449">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-449">
         - Settings</span></span><br><span data-ttu-id="b6e77-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-450">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-451">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-452">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-453">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-455">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="b6e77-456">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-456">- Taskpane</span></span><br><span data-ttu-id="b6e77-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e77-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e77-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e77-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-462">-BindingEvents</span></span><br><span data-ttu-id="b6e77-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-463">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-464">customXmlParts</span></span><br><span data-ttu-id="b6e77-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-465">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-466">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-466">
         - File</span></span><br><span data-ttu-id="b6e77-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-468">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-469">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-472">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-473">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-473">
         - Selection</span></span><br><span data-ttu-id="b6e77-474">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-474">
         - Settings</span></span><br><span data-ttu-id="b6e77-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-475">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-476">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-477">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-478">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-480">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="b6e77-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b6e77-481">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="b6e77-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e77-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e77-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e77-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e77-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e77-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-486">-BindingEvents</span></span><br><span data-ttu-id="b6e77-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-487">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-488">customXmlParts</span></span><br><span data-ttu-id="b6e77-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-489">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-490">
         - File</span></span><br><span data-ttu-id="b6e77-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-492">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-493">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-496">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-497">
         - Selection</span></span><br><span data-ttu-id="b6e77-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-498">
         - Settings</span></span><br><span data-ttu-id="b6e77-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-499">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-500">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-501">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-502">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-504">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e77-505">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-505">- Taskpane</span></span><br><span data-ttu-id="b6e77-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e77-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e77-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e77-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e77-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e77-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-511">-BindingEvents</span></span><br><span data-ttu-id="b6e77-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-512">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-513">customXmlParts</span></span><br><span data-ttu-id="b6e77-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-514">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-515">
         - File</span></span><br><span data-ttu-id="b6e77-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-517">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-518">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-521">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-522">
         - Selection</span></span><br><span data-ttu-id="b6e77-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-523">
         - Settings</span></span><br><span data-ttu-id="b6e77-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-524">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-525">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-526">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-527">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-529">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="b6e77-530">- Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-530">- Taskpane</span></span><br><span data-ttu-id="b6e77-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b6e77-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b6e77-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b6e77-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e77-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e77-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-536">-BindingEvents</span></span><br><span data-ttu-id="b6e77-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-537">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b6e77-538">customXmlParts</span></span><br><span data-ttu-id="b6e77-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-539">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-540">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-540">
         - File</span></span><br><span data-ttu-id="b6e77-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-542">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-543">
         -MatrixBindings</span></span><br><span data-ttu-id="b6e77-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="b6e77-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b6e77-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-546">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-547">
         - Selection</span></span><br><span data-ttu-id="b6e77-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-548">
         - Settings</span></span><br><span data-ttu-id="b6e77-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-549">
         -TableBindings</span></span><br><span data-ttu-id="b6e77-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-550">
         -TableCoercion</span></span><br><span data-ttu-id="b6e77-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b6e77-551">
         -TextBindings</span></span><br><span data-ttu-id="b6e77-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-552">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b6e77-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b6e77-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e77-555">Платформа</span><span class="sxs-lookup"><span data-stu-id="b6e77-555">Platform</span></span></th>
    <th><span data-ttu-id="b6e77-556">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b6e77-556">Extension points</span></span></th>
    <th><span data-ttu-id="b6e77-557">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b6e77-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e77-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e77-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e77-559">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e77-560">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-560">- Content</span></span><br><span data-ttu-id="b6e77-561">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-561">
         - Taskpane</span></span><br><span data-ttu-id="b6e77-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-564">-ActiveView</span></span><br><span data-ttu-id="b6e77-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-565">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-566">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-567">
         - File</span></span><br><span data-ttu-id="b6e77-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-568">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-569">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-570">
         - Selection</span></span><br><span data-ttu-id="b6e77-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-571">
         - Settings</span></span><br><span data-ttu-id="b6e77-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-573">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e77-574">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-574">- Content</span></span><br><span data-ttu-id="b6e77-575">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="b6e77-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b6e77-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b6e77-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-577">-ActiveView</span></span><br><span data-ttu-id="b6e77-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-578">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-579">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-580">
         - File</span></span><br><span data-ttu-id="b6e77-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-581">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-582">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-583">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-583">
         - Selection</span></span><br><span data-ttu-id="b6e77-584">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-584">
         - Settings</span></span><br><span data-ttu-id="b6e77-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-586">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b6e77-587">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-587">- Content</span></span><br><span data-ttu-id="b6e77-588">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-588">
         - Taskpane</span></span><br><span data-ttu-id="b6e77-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-591">-ActiveView</span></span><br><span data-ttu-id="b6e77-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-592">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-593">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-594">
         - File</span></span><br><span data-ttu-id="b6e77-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-595">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-596">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-597">
         - Selection</span></span><br><span data-ttu-id="b6e77-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-598">
         - Settings</span></span><br><span data-ttu-id="b6e77-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-600">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b6e77-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="b6e77-601">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-601">- Content</span></span><br><span data-ttu-id="b6e77-602">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-602">
         - Taskpane</span></span><br><span data-ttu-id="b6e77-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-605">-ActiveView</span></span><br><span data-ttu-id="b6e77-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-606">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-607">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-608">
         - File</span></span><br><span data-ttu-id="b6e77-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-609">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-610">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-611">
         - Selection</span></span><br><span data-ttu-id="b6e77-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-612">
         - Settings</span></span><br><span data-ttu-id="b6e77-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-614">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="b6e77-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b6e77-615">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-615">- Content</span></span><br><span data-ttu-id="b6e77-616">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="b6e77-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="b6e77-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-618">-ActiveView</span></span><br><span data-ttu-id="b6e77-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-619">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-620">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-621">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-621">
         - File</span></span><br><span data-ttu-id="b6e77-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-622">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-623">
         - Selection</span></span><br><span data-ttu-id="b6e77-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-624">
         - Settings</span></span><br><span data-ttu-id="b6e77-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-625">
         -TextCoercion</span></span><br><span data-ttu-id="b6e77-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-627">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b6e77-628">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-628">- Content</span></span><br><span data-ttu-id="b6e77-629">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-629">
         - Taskpane</span></span><br><span data-ttu-id="b6e77-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-632">-ActiveView</span></span><br><span data-ttu-id="b6e77-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-633">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-634">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-635">
         - File</span></span><br><span data-ttu-id="b6e77-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-636">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-637">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-638">
         - Selection</span></span><br><span data-ttu-id="b6e77-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-639">
         - Settings</span></span><br><span data-ttu-id="b6e77-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-641">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b6e77-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="b6e77-642">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-642">- Content</span></span><br><span data-ttu-id="b6e77-643">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-643">
         - Taskpane</span></span><br><span data-ttu-id="b6e77-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b6e77-646">-ActiveView</span></span><br><span data-ttu-id="b6e77-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-647">
         -CompressedFile</span></span><br><span data-ttu-id="b6e77-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-648">
         -DocumentEvents</span></span><br><span data-ttu-id="b6e77-649">
         - File</span><span class="sxs-lookup"><span data-stu-id="b6e77-649">
         - File</span></span><br><span data-ttu-id="b6e77-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-650">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b6e77-651">
         -PdfFile</span></span><br><span data-ttu-id="b6e77-652">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b6e77-652">
         - Selection</span></span><br><span data-ttu-id="b6e77-653">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-653">
         - Settings</span></span><br><span data-ttu-id="b6e77-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="b6e77-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="b6e77-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b6e77-656">Платформа</span><span class="sxs-lookup"><span data-stu-id="b6e77-656">Platform</span></span></th>
    <th><span data-ttu-id="b6e77-657">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b6e77-657">Extension points</span></span></th>
    <th><span data-ttu-id="b6e77-658">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b6e77-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="b6e77-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="b6e77-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="b6e77-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="b6e77-660">Office Online</span></span></td>
    <td> <span data-ttu-id="b6e77-661">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="b6e77-661">- Content</span></span><br><span data-ttu-id="b6e77-662">
         - Панель задач</span><span class="sxs-lookup"><span data-stu-id="b6e77-662">
         - Taskpane</span></span><br><span data-ttu-id="b6e77-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b6e77-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b6e77-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b6e77-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b6e77-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b6e77-666">-DocumentEvents</span></span><br><span data-ttu-id="b6e77-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="b6e77-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-668">
         -ImageCoercion</span></span><br><span data-ttu-id="b6e77-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b6e77-669">
         - Settings</span></span><br><span data-ttu-id="b6e77-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b6e77-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b6e77-671">См. также</span><span class="sxs-lookup"><span data-stu-id="b6e77-671">See also</span></span>

- [<span data-ttu-id="b6e77-672">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b6e77-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b6e77-673">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b6e77-673">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="b6e77-674">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="b6e77-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="b6e77-675">Ссылка на API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="b6e77-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
