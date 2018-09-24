---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы требований для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 09/19/2018
ms.openlocfilehash: 09fb72c88bd0496c413f94b7ba4149192380d664
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967706"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="6ae3c-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6ae3c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="6ae3c-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="6ae3c-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="6ae3c-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="6ae3c-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="6ae3c-106">Символ \* (звездочка) в ячейке таблицы указывает, что поддержка скоро появится.</span><span class="sxs-lookup"><span data-stu-id="6ae3c-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="6ae3c-107">С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="6ae3c-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="6ae3c-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="6ae3c-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="6ae3c-110">Excel</span><span class="sxs-lookup"><span data-stu-id="6ae3c-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6ae3c-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="6ae3c-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6ae3c-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6ae3c-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6ae3c-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6ae3c-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6ae3c-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="6ae3c-115">Office Online</span></span></td>
    <td> <span data-ttu-id="6ae3c-116">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-116">- Taskpane</span></span><br><span data-ttu-id="6ae3c-117">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-117">
        - Content</span></span><br><span data-ttu-id="6ae3c-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="6ae3c-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6ae3c-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6ae3c-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6ae3c-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6ae3c-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="6ae3c-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6ae3c-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-127">
        -BindingEvents</span></span><br><span data-ttu-id="6ae3c-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-128">
        -CompressedFile</span></span><br><span data-ttu-id="6ae3c-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-129">
        -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-130">
        - File</span></span><br><span data-ttu-id="6ae3c-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-131">
        -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-133">
        - Selection</span></span><br><span data-ttu-id="6ae3c-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-134">
        - Settings</span></span><br><span data-ttu-id="6ae3c-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-135">
        -TableBindings</span></span><br><span data-ttu-id="6ae3c-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-136">
        -TableCoercion</span></span><br><span data-ttu-id="6ae3c-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-137">
        -TextBindings</span></span><br><span data-ttu-id="6ae3c-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-139">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="6ae3c-140">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-140">
        - Taskpane</span></span><br><span data-ttu-id="6ae3c-141">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="6ae3c-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6ae3c-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-143">
        -BindingEvents</span></span><br><span data-ttu-id="6ae3c-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-144">
        -CompressedFile</span></span><br><span data-ttu-id="6ae3c-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-145">
        -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-146">
        - File</span></span><br><span data-ttu-id="6ae3c-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-147">
        -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-148">
        -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-150">
        - Selection</span></span><br><span data-ttu-id="6ae3c-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-151">
        - Settings</span></span><br><span data-ttu-id="6ae3c-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-152">
        -TableBindings</span></span><br><span data-ttu-id="6ae3c-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-153">
        -TableCoercion</span></span><br><span data-ttu-id="6ae3c-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-154">
        -TextBindings</span></span><br><span data-ttu-id="6ae3c-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-156">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="6ae3c-157">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-157">- Taskpane</span></span><br><span data-ttu-id="6ae3c-158">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-158">
        - Content</span></span><br><span data-ttu-id="6ae3c-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6ae3c-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6ae3c-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6ae3c-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6ae3c-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="6ae3c-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6ae3c-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-168">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-169">
        -CompressedFile</span></span><br><span data-ttu-id="6ae3c-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-170">
        -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-171">
        - File</span></span><br><span data-ttu-id="6ae3c-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-172">
        -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-173">
        -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-175">
        - Selection</span></span><br><span data-ttu-id="6ae3c-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-176">
        - Settings</span></span><br><span data-ttu-id="6ae3c-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-177">
        -TableBindings</span></span><br><span data-ttu-id="6ae3c-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-178">
        -TableCoercion</span></span><br><span data-ttu-id="6ae3c-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-179">
        -TextBindings</span></span><br><span data-ttu-id="6ae3c-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-181">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="6ae3c-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="6ae3c-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-182">- Taskpane</span></span><br><span data-ttu-id="6ae3c-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-183">
        - Content</span></span></td>
    <td><span data-ttu-id="6ae3c-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6ae3c-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6ae3c-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6ae3c-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="6ae3c-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6ae3c-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-192">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-193">
        -CompressedFile</span></span><br><span data-ttu-id="6ae3c-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-194">
        -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-195">
        - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-195">
        - File</span></span><br><span data-ttu-id="6ae3c-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-196">
        -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-197">
        -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-199">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-199">
        - Selection</span></span><br><span data-ttu-id="6ae3c-200">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-200">
        - Settings</span></span><br><span data-ttu-id="6ae3c-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-201">
        -TableBindings</span></span><br><span data-ttu-id="6ae3c-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-202">
        -TableCoercion</span></span><br><span data-ttu-id="6ae3c-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-203">
        -TextBindings</span></span><br><span data-ttu-id="6ae3c-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-205">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6ae3c-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="6ae3c-206">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-206">- Taskpane</span></span><br><span data-ttu-id="6ae3c-207">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-207">
        - Content</span></span><br><span data-ttu-id="6ae3c-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6ae3c-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6ae3c-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6ae3c-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6ae3c-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="6ae3c-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6ae3c-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-217">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-218">
        -CompressedFile</span></span><br><span data-ttu-id="6ae3c-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-219">
        -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-220">
        - File</span></span><br><span data-ttu-id="6ae3c-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-221">
        -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-222">
        -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-224">
        -PdfFile</span></span><br><span data-ttu-id="6ae3c-225">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-225">
        - Selection</span></span><br><span data-ttu-id="6ae3c-226">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-226">
        - Settings</span></span><br><span data-ttu-id="6ae3c-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-227">
        -TableBindings</span></span><br><span data-ttu-id="6ae3c-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-228">
        -TableCoercion</span></span><br><span data-ttu-id="6ae3c-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-229">
        -TextBindings</span></span><br><span data-ttu-id="6ae3c-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="6ae3c-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="6ae3c-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6ae3c-232">Платформа</span><span class="sxs-lookup"><span data-stu-id="6ae3c-232">Platform</span></span></th>
    <th><span data-ttu-id="6ae3c-233">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6ae3c-233">Extension points</span></span></th>
    <th><span data-ttu-id="6ae3c-234">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6ae3c-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="6ae3c-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="6ae3c-236">Office Online</span></span></td>
    <td> <span data-ttu-id="6ae3c-237">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-237">- Mail Read</span></span><br><span data-ttu-id="6ae3c-238">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-238">
      - Mail Compose</span></span><br><span data-ttu-id="6ae3c-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6ae3c-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6ae3c-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6ae3c-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6ae3c-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6ae3c-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6ae3c-246">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6ae3c-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-247">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6ae3c-248">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-248">- Mail Read</span></span><br><span data-ttu-id="6ae3c-249">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-249">
      - Mail Compose</span></span><br><span data-ttu-id="6ae3c-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6ae3c-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6ae3c-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6ae3c-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="6ae3c-255">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6ae3c-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-256">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6ae3c-257">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-257">- Mail Read</span></span><br><span data-ttu-id="6ae3c-258">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-258">
      - Mail Compose</span></span><br><span data-ttu-id="6ae3c-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6ae3c-260">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="6ae3c-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6ae3c-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6ae3c-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6ae3c-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6ae3c-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6ae3c-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6ae3c-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6ae3c-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="6ae3c-268">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6ae3c-268">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-269">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="6ae3c-269">Office for iOS</span></span></td>
    <td> <span data-ttu-id="6ae3c-270">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-270">- Mail Read</span></span><br><span data-ttu-id="6ae3c-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6ae3c-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6ae3c-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6ae3c-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6ae3c-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6ae3c-277">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6ae3c-277">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-278">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6ae3c-278">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6ae3c-279">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-279">- Mail Read</span></span><br><span data-ttu-id="6ae3c-280">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-280">
      - Mail Compose</span></span><br><span data-ttu-id="6ae3c-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6ae3c-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6ae3c-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6ae3c-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6ae3c-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6ae3c-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6ae3c-288">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6ae3c-288">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-289">Office для Android</span><span class="sxs-lookup"><span data-stu-id="6ae3c-289">Office for Android</span></span></td>
    <td> <span data-ttu-id="6ae3c-290">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6ae3c-290">- Mail Read</span></span><br><span data-ttu-id="6ae3c-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6ae3c-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6ae3c-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6ae3c-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6ae3c-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6ae3c-297">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6ae3c-297">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="6ae3c-298">Word</span><span class="sxs-lookup"><span data-stu-id="6ae3c-298">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6ae3c-299">Платформа</span><span class="sxs-lookup"><span data-stu-id="6ae3c-299">Platform</span></span></th>
    <th><span data-ttu-id="6ae3c-300">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6ae3c-300">Extension points</span></span></th>
    <th><span data-ttu-id="6ae3c-301">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6ae3c-301">API requirement sets</span></span></th>
    <th><span data-ttu-id="6ae3c-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-303">Office Online</span><span class="sxs-lookup"><span data-stu-id="6ae3c-303">Office Online</span></span></td>
    <td> <span data-ttu-id="6ae3c-304">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-304">- Taskpane</span></span><br><span data-ttu-id="6ae3c-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-310">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-310">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-311">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6ae3c-311">customXmlParts</span></span><br><span data-ttu-id="6ae3c-312">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-312">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-313">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-313">
         - File</span></span><br><span data-ttu-id="6ae3c-314">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-314">
         -HtmlCoercion</span></span><br><span data-ttu-id="6ae3c-315">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-315">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-316">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-316">
         -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-317">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-317">
         -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-318">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-318">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6ae3c-319">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-319">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-320">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-320">
         - Selection</span></span><br><span data-ttu-id="6ae3c-321">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-321">
         - Settings</span></span><br><span data-ttu-id="6ae3c-322">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-322">
         -TableBindings</span></span><br><span data-ttu-id="6ae3c-323">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-323">
         -TableCoercion</span></span><br><span data-ttu-id="6ae3c-324">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-324">
         -TextBindings</span></span><br><span data-ttu-id="6ae3c-325">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-325">
         -TextCoercion</span></span><br><span data-ttu-id="6ae3c-326">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-326">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-327">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6ae3c-328">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-328">- Taskpane</span></span></td>
    <td> <span data-ttu-id="6ae3c-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-330">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-330">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-331">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-331">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-332">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6ae3c-332">customXmlParts</span></span><br><span data-ttu-id="6ae3c-333">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-333">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-334">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-334">
         - File</span></span><br><span data-ttu-id="6ae3c-335">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-335">
         -HtmlCoercion</span></span><br><span data-ttu-id="6ae3c-336">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-336">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-337">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-337">
         -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-338">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-338">
         -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-339">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-339">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6ae3c-340">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-340">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-341">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-341">
         - Selection</span></span><br><span data-ttu-id="6ae3c-342">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-342">
         - Settings</span></span><br><span data-ttu-id="6ae3c-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-343">
         -TableBindings</span></span><br><span data-ttu-id="6ae3c-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-344">
         -TableCoercion</span></span><br><span data-ttu-id="6ae3c-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-345">
         -TextBindings</span></span><br><span data-ttu-id="6ae3c-346">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-346">
         -TextCoercion</span></span><br><span data-ttu-id="6ae3c-347">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-347">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-348">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-348">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6ae3c-349">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-349">- Taskpane</span></span><br><span data-ttu-id="6ae3c-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-355">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-355">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-356">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-356">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-357">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6ae3c-357">customXmlParts</span></span><br><span data-ttu-id="6ae3c-358">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-358">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-359">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-359">
         - File</span></span><br><span data-ttu-id="6ae3c-360">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-360">
         -HtmlCoercion</span></span><br><span data-ttu-id="6ae3c-361">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-361">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-362">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-362">
         -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-364">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-364">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6ae3c-365">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-365">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-366">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-366">
         - Selection</span></span><br><span data-ttu-id="6ae3c-367">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-367">
         - Settings</span></span><br><span data-ttu-id="6ae3c-368">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-368">
         -TableBindings</span></span><br><span data-ttu-id="6ae3c-369">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-369">
         -TableCoercion</span></span><br><span data-ttu-id="6ae3c-370">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-370">
         -TextBindings</span></span><br><span data-ttu-id="6ae3c-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-371">
         -TextCoercion</span></span><br><span data-ttu-id="6ae3c-372">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-372">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-373">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="6ae3c-373">Office for iOS</span></span></td>
    <td> <span data-ttu-id="6ae3c-374">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-374">- Taskpane</span></span></td>
    <td> <span data-ttu-id="6ae3c-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6ae3c-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6ae3c-379">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-379">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-380">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-380">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-381">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6ae3c-381">customXmlParts</span></span><br><span data-ttu-id="6ae3c-382">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-382">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-383">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-383">
         - File</span></span><br><span data-ttu-id="6ae3c-384">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-384">
         -HtmlCoercion</span></span><br><span data-ttu-id="6ae3c-385">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-385">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-386">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-386">
         -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-387">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-387">
         -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-388">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-388">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6ae3c-389">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-389">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-390">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-390">
         - Selection</span></span><br><span data-ttu-id="6ae3c-391">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-391">
         - Settings</span></span><br><span data-ttu-id="6ae3c-392">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-392">
         -TableBindings</span></span><br><span data-ttu-id="6ae3c-393">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-393">
         -TableCoercion</span></span><br><span data-ttu-id="6ae3c-394">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-394">
         -TextBindings</span></span><br><span data-ttu-id="6ae3c-395">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-395">
         -TextCoercion</span></span><br><span data-ttu-id="6ae3c-396">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-396">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-397">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6ae3c-397">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6ae3c-398">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-398">- Taskpane</span></span><br><span data-ttu-id="6ae3c-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6ae3c-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6ae3c-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6ae3c-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6ae3c-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-404">-BindingEvents</span></span><br><span data-ttu-id="6ae3c-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-405">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6ae3c-406">customXmlParts</span></span><br><span data-ttu-id="6ae3c-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-407">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-408">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-408">
         - File</span></span><br><span data-ttu-id="6ae3c-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="6ae3c-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-410">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-411">
         -MatrixBindings</span></span><br><span data-ttu-id="6ae3c-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="6ae3c-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="6ae3c-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-414">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-415">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-415">
         - Selection</span></span><br><span data-ttu-id="6ae3c-416">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-416">
         - Settings</span></span><br><span data-ttu-id="6ae3c-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-417">
         -TableBindings</span></span><br><span data-ttu-id="6ae3c-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-418">
         -TableCoercion</span></span><br><span data-ttu-id="6ae3c-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-419">
         -TextBindings</span></span><br><span data-ttu-id="6ae3c-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-420">
         -TextCoercion</span></span><br><span data-ttu-id="6ae3c-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-421">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="6ae3c-422">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6ae3c-422">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6ae3c-423">Платформа</span><span class="sxs-lookup"><span data-stu-id="6ae3c-423">Platform</span></span></th>
    <th><span data-ttu-id="6ae3c-424">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6ae3c-424">Extension points</span></span></th>
    <th><span data-ttu-id="6ae3c-425">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6ae3c-425">API requirement sets</span></span></th>
    <th><span data-ttu-id="6ae3c-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-427">Office Online</span><span class="sxs-lookup"><span data-stu-id="6ae3c-427">Office Online</span></span></td>
    <td> <span data-ttu-id="6ae3c-428">- Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-428">- Content</span></span><br><span data-ttu-id="6ae3c-429">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-429">
         - Taskpane</span></span><br><span data-ttu-id="6ae3c-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6ae3c-432">-ActiveView</span></span><br><span data-ttu-id="6ae3c-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-433">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-434">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-435">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-435">
         - File</span></span><br><span data-ttu-id="6ae3c-436">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-436">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-437">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-437">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-438">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-438">
         - Selection</span></span><br><span data-ttu-id="6ae3c-439">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-439">
         - Settings</span></span><br><span data-ttu-id="6ae3c-440">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-440">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-441">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-441">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6ae3c-442">- Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-442">- Content</span></span><br><span data-ttu-id="6ae3c-443">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-443">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="6ae3c-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6ae3c-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6ae3c-445">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6ae3c-445">-ActiveView</span></span><br><span data-ttu-id="6ae3c-446">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-446">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-447">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-447">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-448">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-448">
         - File</span></span><br><span data-ttu-id="6ae3c-449">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-449">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-450">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-451">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-451">
         - Selection</span></span><br><span data-ttu-id="6ae3c-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-452">
         - Settings</span></span><br><span data-ttu-id="6ae3c-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-453">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-454">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6ae3c-454">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6ae3c-455">- Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-455">- Content</span></span><br><span data-ttu-id="6ae3c-456">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-456">
         - Taskpane</span></span><br><span data-ttu-id="6ae3c-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-459">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6ae3c-459">-ActiveView</span></span><br><span data-ttu-id="6ae3c-460">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-460">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-461">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-461">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-462">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-462">
         - File</span></span><br><span data-ttu-id="6ae3c-463">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-463">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-464">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-465">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-465">
         - Selection</span></span><br><span data-ttu-id="6ae3c-466">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-466">
         - Settings</span></span><br><span data-ttu-id="6ae3c-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-467">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-468">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="6ae3c-468">Office for iOS</span></span></td>
    <td> <span data-ttu-id="6ae3c-469">- Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-469">- Content</span></span><br><span data-ttu-id="6ae3c-470">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-470">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="6ae3c-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="6ae3c-472">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6ae3c-472">-ActiveView</span></span><br><span data-ttu-id="6ae3c-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-473">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-474">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-474">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-475">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-475">
         - File</span></span><br><span data-ttu-id="6ae3c-476">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-476">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-477">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-477">
         - Selection</span></span><br><span data-ttu-id="6ae3c-478">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-478">
         - Settings</span></span><br><span data-ttu-id="6ae3c-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-479">
         -TextCoercion</span></span><br><span data-ttu-id="6ae3c-480">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-480">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-481">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6ae3c-481">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6ae3c-482">- Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-482">- Content</span></span><br><span data-ttu-id="6ae3c-483">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-483">
         - Taskpane</span></span><br><span data-ttu-id="6ae3c-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-486">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6ae3c-486">-ActiveView</span></span><br><span data-ttu-id="6ae3c-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-487">
         -CompressedFile</span></span><br><span data-ttu-id="6ae3c-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-488">
         -DocumentEvents</span></span><br><span data-ttu-id="6ae3c-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="6ae3c-489">
         - File</span></span><br><span data-ttu-id="6ae3c-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-490">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-491">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6ae3c-491">
         -PdfFile</span></span><br><span data-ttu-id="6ae3c-492">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6ae3c-492">
         - Selection</span></span><br><span data-ttu-id="6ae3c-493">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-493">
         - Settings</span></span><br><span data-ttu-id="6ae3c-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-494">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="6ae3c-495">OneNote</span><span class="sxs-lookup"><span data-stu-id="6ae3c-495">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6ae3c-496">Платформа</span><span class="sxs-lookup"><span data-stu-id="6ae3c-496">Platform</span></span></th>
    <th><span data-ttu-id="6ae3c-497">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6ae3c-497">Extension points</span></span></th>
    <th><span data-ttu-id="6ae3c-498">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6ae3c-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="6ae3c-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="6ae3c-500">Office Online</span><span class="sxs-lookup"><span data-stu-id="6ae3c-500">Office Online</span></span></td>
    <td> <span data-ttu-id="6ae3c-501">- Контент</span><span class="sxs-lookup"><span data-stu-id="6ae3c-501">- Content</span></span><br><span data-ttu-id="6ae3c-502">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6ae3c-502">
         - Taskpane</span></span><br><span data-ttu-id="6ae3c-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="6ae3c-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6ae3c-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6ae3c-506">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6ae3c-506">-DocumentEvents</span></span><br><span data-ttu-id="6ae3c-507">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-507">
         -HtmlCoercion</span></span><br><span data-ttu-id="6ae3c-508">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-508">
         -ImageCoercion</span></span><br><span data-ttu-id="6ae3c-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6ae3c-509">
         - Settings</span></span><br><span data-ttu-id="6ae3c-510">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6ae3c-510">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="6ae3c-511">См. также</span><span class="sxs-lookup"><span data-stu-id="6ae3c-511">See also</span></span>

- [<span data-ttu-id="6ae3c-512">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6ae3c-512">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="6ae3c-513">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6ae3c-513">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="6ae3c-514">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="6ae3c-514">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="6ae3c-515">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="6ae3c-515">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
