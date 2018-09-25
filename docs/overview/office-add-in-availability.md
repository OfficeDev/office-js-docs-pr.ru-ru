---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы требований для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 09/24/2018
ms.openlocfilehash: b06602e35ec906866ad16d667036a4cbaff2d89e
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985825"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="bf39e-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bf39e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="bf39e-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="bf39e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="bf39e-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="bf39e-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="bf39e-106">Символ \* (звездочка) в ячейке таблицы указывает, что поддержка скоро появится.</span><span class="sxs-lookup"><span data-stu-id="bf39e-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="bf39e-107">С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="bf39e-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="bf39e-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="bf39e-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="bf39e-110">Excel</span><span class="sxs-lookup"><span data-stu-id="bf39e-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="bf39e-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="bf39e-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="bf39e-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bf39e-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="bf39e-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bf39e-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="bf39e-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bf39e-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="bf39e-115">Office Online</span></span></td>
    <td> <span data-ttu-id="bf39e-116">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-116">- Taskpane</span></span><br><span data-ttu-id="bf39e-117">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-117">
        - Content</span></span><br><span data-ttu-id="bf39e-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="bf39e-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="bf39e-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bf39e-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bf39e-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bf39e-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bf39e-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bf39e-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bf39e-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="bf39e-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-127">
        -BindingEvents</span></span><br><span data-ttu-id="bf39e-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-128">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-129">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-130">
        - File</span></span><br><span data-ttu-id="bf39e-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-131">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-133">
        - Selection</span></span><br><span data-ttu-id="bf39e-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-134">
        - Settings</span></span><br><span data-ttu-id="bf39e-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-135">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-136">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-137">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-139">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="bf39e-140">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-140">
        - Taskpane</span></span><br><span data-ttu-id="bf39e-141">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="bf39e-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-143">
        -BindingEvents</span></span><br><span data-ttu-id="bf39e-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-144">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-145">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-146">
        - File</span></span><br><span data-ttu-id="bf39e-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-147">
        -ImageCoercion</span></span><br><span data-ttu-id="bf39e-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-148">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-150">
        - Selection</span></span><br><span data-ttu-id="bf39e-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-151">
        - Settings</span></span><br><span data-ttu-id="bf39e-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-152">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-153">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-154">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-156">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="bf39e-157">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-157">- Taskpane</span></span><br><span data-ttu-id="bf39e-158">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-158">
        - Content</span></span><br><span data-ttu-id="bf39e-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bf39e-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bf39e-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bf39e-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bf39e-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bf39e-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bf39e-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bf39e-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="bf39e-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-168">-BindingEvents</span></span><br><span data-ttu-id="bf39e-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-169">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-170">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-171">
        - File</span></span><br><span data-ttu-id="bf39e-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-172">
        -ImageCoercion</span></span><br><span data-ttu-id="bf39e-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-173">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-175">
        - Selection</span></span><br><span data-ttu-id="bf39e-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-176">
        - Settings</span></span><br><span data-ttu-id="bf39e-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-177">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-178">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-179">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-181">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-181">Office for Windows</span></span></td>
    <td><span data-ttu-id="bf39e-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-182">- Taskpane</span></span><br><span data-ttu-id="bf39e-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-183">
        - Content</span></span><br><span data-ttu-id="bf39e-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bf39e-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bf39e-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bf39e-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bf39e-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bf39e-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bf39e-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bf39e-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="bf39e-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-193">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-193">-BindingEvents</span></span><br><span data-ttu-id="bf39e-194">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-194">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-195">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-195">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-196">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-196">
        - File</span></span><br><span data-ttu-id="bf39e-197">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-197">
        -ImageCoercion</span></span><br><span data-ttu-id="bf39e-198">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-198">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-199">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-199">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-200">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-200">
        - Selection</span></span><br><span data-ttu-id="bf39e-201">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-201">
        - Settings</span></span><br><span data-ttu-id="bf39e-202">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-202">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-203">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-203">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-204">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-204">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-205">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-205">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-206">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="bf39e-206">Office for iOS</span></span></td>
    <td><span data-ttu-id="bf39e-207">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-207">- Taskpane</span></span><br><span data-ttu-id="bf39e-208">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-208">
        - Content</span></span></td>
    <td><span data-ttu-id="bf39e-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bf39e-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bf39e-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bf39e-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bf39e-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bf39e-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bf39e-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="bf39e-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-217">-BindingEvents</span></span><br><span data-ttu-id="bf39e-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-218">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-219">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-220">
        - File</span></span><br><span data-ttu-id="bf39e-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-221">
        -ImageCoercion</span></span><br><span data-ttu-id="bf39e-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-222">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-224">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-224">
        - Selection</span></span><br><span data-ttu-id="bf39e-225">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-225">
        - Settings</span></span><br><span data-ttu-id="bf39e-226">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-226">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-227">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-227">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-228">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-228">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-229">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-229">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-230">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-230">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="bf39e-231">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-231">- Taskpane</span></span><br><span data-ttu-id="bf39e-232">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-232">
        - Content</span></span><br><span data-ttu-id="bf39e-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bf39e-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bf39e-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bf39e-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bf39e-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bf39e-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bf39e-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bf39e-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-240">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="bf39e-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-242">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-242">-BindingEvents</span></span><br><span data-ttu-id="bf39e-243">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-243">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-244">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-244">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-245">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-245">
        - File</span></span><br><span data-ttu-id="bf39e-246">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-246">
        -ImageCoercion</span></span><br><span data-ttu-id="bf39e-247">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-247">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-248">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-248">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-249">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-249">
        -PdfFile</span></span><br><span data-ttu-id="bf39e-250">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-250">
        - Selection</span></span><br><span data-ttu-id="bf39e-251">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-251">
        - Settings</span></span><br><span data-ttu-id="bf39e-252">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-252">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-253">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-253">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-254">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-254">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-255">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-255">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-256">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-256">Office for Mac</span></span></td>
    <td><span data-ttu-id="bf39e-257">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-257">- Taskpane</span></span><br><span data-ttu-id="bf39e-258">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-258">
        - Content</span></span><br><span data-ttu-id="bf39e-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bf39e-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bf39e-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bf39e-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bf39e-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bf39e-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bf39e-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bf39e-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-266">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="bf39e-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bf39e-268">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-268">-BindingEvents</span></span><br><span data-ttu-id="bf39e-269">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-269">
        -CompressedFile</span></span><br><span data-ttu-id="bf39e-270">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-270">
        -DocumentEvents</span></span><br><span data-ttu-id="bf39e-271">
        - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-271">
        - File</span></span><br><span data-ttu-id="bf39e-272">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-272">
        -ImageCoercion</span></span><br><span data-ttu-id="bf39e-273">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-273">
        -MatrixBindings</span></span><br><span data-ttu-id="bf39e-274">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-274">
        -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-275">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-275">
        -PdfFile</span></span><br><span data-ttu-id="bf39e-276">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-276">
        - Selection</span></span><br><span data-ttu-id="bf39e-277">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-277">
        - Settings</span></span><br><span data-ttu-id="bf39e-278">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-278">
        -TableBindings</span></span><br><span data-ttu-id="bf39e-279">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-279">
        -TableCoercion</span></span><br><span data-ttu-id="bf39e-280">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-280">
        -TextBindings</span></span><br><span data-ttu-id="bf39e-281">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-281">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="bf39e-282">Outlook</span><span class="sxs-lookup"><span data-stu-id="bf39e-282">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bf39e-283">Платформа</span><span class="sxs-lookup"><span data-stu-id="bf39e-283">Platform</span></span></th>
    <th><span data-ttu-id="bf39e-284">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bf39e-284">Extension points</span></span></th>
    <th><span data-ttu-id="bf39e-285">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bf39e-285">API requirement sets</span></span></th>
    <th><span data-ttu-id="bf39e-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bf39e-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-287">Office Online</span><span class="sxs-lookup"><span data-stu-id="bf39e-287">Office Online</span></span></td>
    <td> <span data-ttu-id="bf39e-288">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-288">- Mail Read</span></span><br><span data-ttu-id="bf39e-289">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-289">
      - Mail Compose</span></span><br><span data-ttu-id="bf39e-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bf39e-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bf39e-297">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-297">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-298">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-298">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-299">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-299">- Mail Read</span></span><br><span data-ttu-id="bf39e-300">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-300">
      - Mail Compose</span></span><br><span data-ttu-id="bf39e-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="bf39e-306">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-306">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-307">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-307">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-308">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-308">- Mail Read</span></span><br><span data-ttu-id="bf39e-309">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-309">
      - Mail Compose</span></span><br><span data-ttu-id="bf39e-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bf39e-311">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="bf39e-311">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bf39e-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bf39e-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bf39e-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bf39e-319">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-319">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-320">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-320">Office for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-321">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-321">- Mail Read</span></span><br><span data-ttu-id="bf39e-322">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-322">
      - Mail Compose</span></span><br><span data-ttu-id="bf39e-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bf39e-324">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="bf39e-324">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bf39e-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bf39e-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bf39e-331">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-331">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-332">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="bf39e-332">Office for iOS</span></span></td>
    <td> <span data-ttu-id="bf39e-333">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-333">- Mail Read</span></span><br><span data-ttu-id="bf39e-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bf39e-340">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-341">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-341">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bf39e-342">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-342">- Mail Read</span></span><br><span data-ttu-id="bf39e-343">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-343">
      - Mail Compose</span></span><br><span data-ttu-id="bf39e-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bf39e-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bf39e-351">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-352">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-352">Office for Mac</span></span></td>
    <td> <span data-ttu-id="bf39e-353">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-353">- Mail Read</span></span><br><span data-ttu-id="bf39e-354">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-354">
      - Mail Compose</span></span><br><span data-ttu-id="bf39e-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bf39e-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bf39e-362">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-362">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-363">Office для Android</span><span class="sxs-lookup"><span data-stu-id="bf39e-363">Office for Android</span></span></td>
    <td> <span data-ttu-id="bf39e-364">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bf39e-364">- Mail Read</span></span><br><span data-ttu-id="bf39e-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bf39e-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bf39e-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bf39e-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bf39e-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bf39e-371">Недоступна</span><span class="sxs-lookup"><span data-stu-id="bf39e-371">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="bf39e-372">Word</span><span class="sxs-lookup"><span data-stu-id="bf39e-372">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bf39e-373">Платформа</span><span class="sxs-lookup"><span data-stu-id="bf39e-373">Platform</span></span></th>
    <th><span data-ttu-id="bf39e-374">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bf39e-374">Extension points</span></span></th>
    <th><span data-ttu-id="bf39e-375">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bf39e-375">API requirement sets</span></span></th>
    <th><span data-ttu-id="bf39e-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bf39e-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-377">Office Online</span><span class="sxs-lookup"><span data-stu-id="bf39e-377">Office Online</span></span></td>
    <td> <span data-ttu-id="bf39e-378">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-378">- Taskpane</span></span><br><span data-ttu-id="bf39e-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bf39e-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bf39e-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bf39e-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-384">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-384">-BindingEvents</span></span><br><span data-ttu-id="bf39e-385">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-385">customXmlParts</span></span><br><span data-ttu-id="bf39e-386">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-386">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-387">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-387">
         - File</span></span><br><span data-ttu-id="bf39e-388">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-388">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-389">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-389">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-390">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-390">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-391">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-391">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-392">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-392">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-393">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-393">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-394">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-394">
         - Selection</span></span><br><span data-ttu-id="bf39e-395">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-395">
         - Settings</span></span><br><span data-ttu-id="bf39e-396">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-396">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-397">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-397">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-398">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-398">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-399">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-399">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-400">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-400">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-401">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-401">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-402">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-402">- Taskpane</span></span></td>
    <td> <span data-ttu-id="bf39e-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-404">-BindingEvents</span></span><br><span data-ttu-id="bf39e-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-405">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-406">customXmlParts</span></span><br><span data-ttu-id="bf39e-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-407">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-408">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-408">
         - File</span></span><br><span data-ttu-id="bf39e-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-410">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-411">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-414">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-415">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-415">
         - Selection</span></span><br><span data-ttu-id="bf39e-416">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-416">
         - Settings</span></span><br><span data-ttu-id="bf39e-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-417">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-418">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-419">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-420">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-421">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-422">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-422">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-423">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-423">- Taskpane</span></span><br><span data-ttu-id="bf39e-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bf39e-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bf39e-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bf39e-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-429">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-429">-BindingEvents</span></span><br><span data-ttu-id="bf39e-430">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-430">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-431">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-431">customXmlParts</span></span><br><span data-ttu-id="bf39e-432">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-432">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-433">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-433">
         - File</span></span><br><span data-ttu-id="bf39e-434">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-434">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-435">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-436">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-436">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-437">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-437">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-438">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-438">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-439">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-439">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-440">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-440">
         - Selection</span></span><br><span data-ttu-id="bf39e-441">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-441">
         - Settings</span></span><br><span data-ttu-id="bf39e-442">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-442">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-443">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-443">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-444">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-444">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-445">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-445">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-446">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-446">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-447">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-447">Office for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-448">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-448">- Taskpane</span></span><br><span data-ttu-id="bf39e-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bf39e-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bf39e-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bf39e-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-454">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-454">-BindingEvents</span></span><br><span data-ttu-id="bf39e-455">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-455">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-456">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-456">customXmlParts</span></span><br><span data-ttu-id="bf39e-457">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-457">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-458">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-458">
         - File</span></span><br><span data-ttu-id="bf39e-459">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-459">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-460">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-460">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-461">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-461">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-462">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-462">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-463">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-463">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-464">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-465">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-465">
         - Selection</span></span><br><span data-ttu-id="bf39e-466">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-466">
         - Settings</span></span><br><span data-ttu-id="bf39e-467">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-467">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-468">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-468">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-469">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-469">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-470">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-470">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-471">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-471">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-472">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="bf39e-472">Office for iOS</span></span></td>
    <td> <span data-ttu-id="bf39e-473">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-473">- Taskpane</span></span></td>
    <td> <span data-ttu-id="bf39e-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bf39e-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bf39e-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bf39e-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bf39e-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bf39e-478">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-478">-BindingEvents</span></span><br><span data-ttu-id="bf39e-479">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-479">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-480">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-480">customXmlParts</span></span><br><span data-ttu-id="bf39e-481">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-481">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-482">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-482">
         - File</span></span><br><span data-ttu-id="bf39e-483">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-483">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-484">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-484">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-485">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-485">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-486">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-486">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-487">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-487">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-488">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-488">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-489">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-489">
         - Selection</span></span><br><span data-ttu-id="bf39e-490">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-490">
         - Settings</span></span><br><span data-ttu-id="bf39e-491">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-491">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-492">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-492">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-493">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-493">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-494">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-495">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-495">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-496">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-496">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bf39e-497">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-497">- Taskpane</span></span><br><span data-ttu-id="bf39e-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bf39e-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bf39e-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bf39e-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bf39e-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bf39e-503">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-503">-BindingEvents</span></span><br><span data-ttu-id="bf39e-504">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-504">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-505">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-505">customXmlParts</span></span><br><span data-ttu-id="bf39e-506">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-506">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-507">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-507">
         - File</span></span><br><span data-ttu-id="bf39e-508">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-508">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-509">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-509">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-510">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-510">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-511">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-511">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-512">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-512">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-513">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-513">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-514">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-514">
         - Selection</span></span><br><span data-ttu-id="bf39e-515">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-515">
         - Settings</span></span><br><span data-ttu-id="bf39e-516">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-516">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-517">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-517">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-518">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-518">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-519">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-519">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-520">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-520">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-521">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-521">Office for Mac</span></span></td>
    <td> <span data-ttu-id="bf39e-522">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-522">- Taskpane</span></span><br><span data-ttu-id="bf39e-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bf39e-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bf39e-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bf39e-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bf39e-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bf39e-528">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-528">-BindingEvents</span></span><br><span data-ttu-id="bf39e-529">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-529">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-530">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bf39e-530">customXmlParts</span></span><br><span data-ttu-id="bf39e-531">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-531">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-532">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-532">
         - File</span></span><br><span data-ttu-id="bf39e-533">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-533">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-534">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-534">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-535">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-535">
         -MatrixBindings</span></span><br><span data-ttu-id="bf39e-536">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-536">
         -MatrixCoercion</span></span><br><span data-ttu-id="bf39e-537">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-537">
         -OoxmlCoercion</span></span><br><span data-ttu-id="bf39e-538">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-538">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-539">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-539">
         - Selection</span></span><br><span data-ttu-id="bf39e-540">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bf39e-540">
         - Settings</span></span><br><span data-ttu-id="bf39e-541">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-541">
         -TableBindings</span></span><br><span data-ttu-id="bf39e-542">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-542">
         -TableCoercion</span></span><br><span data-ttu-id="bf39e-543">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bf39e-543">
         -TextBindings</span></span><br><span data-ttu-id="bf39e-544">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-544">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-545">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-545">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="bf39e-546">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bf39e-546">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bf39e-547">Платформа</span><span class="sxs-lookup"><span data-stu-id="bf39e-547">Platform</span></span></th>
    <th><span data-ttu-id="bf39e-548">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bf39e-548">Extension points</span></span></th>
    <th><span data-ttu-id="bf39e-549">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bf39e-549">API requirement sets</span></span></th>
    <th><span data-ttu-id="bf39e-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bf39e-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-551">Office Online</span><span class="sxs-lookup"><span data-stu-id="bf39e-551">Office Online</span></span></td>
    <td> <span data-ttu-id="bf39e-552">- Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-552">- Content</span></span><br><span data-ttu-id="bf39e-553">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-553">
         - Taskpane</span></span><br><span data-ttu-id="bf39e-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-556">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-556">-ActiveView</span></span><br><span data-ttu-id="bf39e-557">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-557">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-558">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-558">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-559">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-559">
         - File</span></span><br><span data-ttu-id="bf39e-560">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-560">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-561">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-561">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-562">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-562">
         - Selection</span></span><br><span data-ttu-id="bf39e-563">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-563">
         - Settings</span></span><br><span data-ttu-id="bf39e-564">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-564">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-565">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-565">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-566">- Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-566">- Content</span></span><br><span data-ttu-id="bf39e-567">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-567">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="bf39e-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bf39e-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bf39e-569">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-569">-ActiveView</span></span><br><span data-ttu-id="bf39e-570">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-570">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-571">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-572">
         - File</span></span><br><span data-ttu-id="bf39e-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-573">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-574">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-575">
         - Selection</span></span><br><span data-ttu-id="bf39e-576">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-576">
         - Settings</span></span><br><span data-ttu-id="bf39e-577">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-577">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-578">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-578">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-579">- Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-579">- Content</span></span><br><span data-ttu-id="bf39e-580">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-580">
         - Taskpane</span></span><br><span data-ttu-id="bf39e-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-583">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-583">-ActiveView</span></span><br><span data-ttu-id="bf39e-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-584">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-585">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-586">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-586">
         - File</span></span><br><span data-ttu-id="bf39e-587">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-587">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-588">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-588">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-589">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-589">
         - Selection</span></span><br><span data-ttu-id="bf39e-590">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-590">
         - Settings</span></span><br><span data-ttu-id="bf39e-591">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-591">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-592">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bf39e-592">Office for Windows</span></span></td>
    <td> <span data-ttu-id="bf39e-593">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="bf39e-593">- Content</span></span><br><span data-ttu-id="bf39e-594">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-594">
         - Taskpane</span></span><br><span data-ttu-id="bf39e-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-597">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-597">-ActiveView</span></span><br><span data-ttu-id="bf39e-598">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-598">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-599">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-599">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-600">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-600">
         - File</span></span><br><span data-ttu-id="bf39e-601">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-601">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-602">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-602">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-603">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-603">
         - Selection</span></span><br><span data-ttu-id="bf39e-604">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-604">
         - Settings</span></span><br><span data-ttu-id="bf39e-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-605">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-606">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="bf39e-606">Office for iOS</span></span></td>
    <td> <span data-ttu-id="bf39e-607">- Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-607">- Content</span></span><br><span data-ttu-id="bf39e-608">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-608">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="bf39e-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="bf39e-610">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-610">-ActiveView</span></span><br><span data-ttu-id="bf39e-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-611">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-612">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-613">
         - File</span></span><br><span data-ttu-id="bf39e-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-614">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-615">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-615">
         - Selection</span></span><br><span data-ttu-id="bf39e-616">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-616">
         - Settings</span></span><br><span data-ttu-id="bf39e-617">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-617">
         -TextCoercion</span></span><br><span data-ttu-id="bf39e-618">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-618">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-619">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-619">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bf39e-620">- Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-620">- Content</span></span><br><span data-ttu-id="bf39e-621">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-621">
         - Taskpane</span></span><br><span data-ttu-id="bf39e-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-624">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-624">-ActiveView</span></span><br><span data-ttu-id="bf39e-625">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-625">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-626">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-627">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-627">
         - File</span></span><br><span data-ttu-id="bf39e-628">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-628">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-629">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-629">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-630">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-630">
         - Selection</span></span><br><span data-ttu-id="bf39e-631">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-631">
         - Settings</span></span><br><span data-ttu-id="bf39e-632">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-632">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-633">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bf39e-633">Office for Mac</span></span></td>
    <td> <span data-ttu-id="bf39e-634">- Содержимое</span><span class="sxs-lookup"><span data-stu-id="bf39e-634">- Content</span></span><br><span data-ttu-id="bf39e-635">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-635">
         - Taskpane</span></span><br><span data-ttu-id="bf39e-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-638">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bf39e-638">-ActiveView</span></span><br><span data-ttu-id="bf39e-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-639">
         -CompressedFile</span></span><br><span data-ttu-id="bf39e-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-640">
         -DocumentEvents</span></span><br><span data-ttu-id="bf39e-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="bf39e-641">
         - File</span></span><br><span data-ttu-id="bf39e-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-642">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-643">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bf39e-643">
         -PdfFile</span></span><br><span data-ttu-id="bf39e-644">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bf39e-644">
         - Selection</span></span><br><span data-ttu-id="bf39e-645">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-645">
         - Settings</span></span><br><span data-ttu-id="bf39e-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-646">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="bf39e-647">OneNote</span><span class="sxs-lookup"><span data-stu-id="bf39e-647">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bf39e-648">Платформа</span><span class="sxs-lookup"><span data-stu-id="bf39e-648">Platform</span></span></th>
    <th><span data-ttu-id="bf39e-649">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bf39e-649">Extension points</span></span></th>
    <th><span data-ttu-id="bf39e-650">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bf39e-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="bf39e-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bf39e-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="bf39e-652">Office Online</span><span class="sxs-lookup"><span data-stu-id="bf39e-652">Office Online</span></span></td>
    <td> <span data-ttu-id="bf39e-653">- Контент</span><span class="sxs-lookup"><span data-stu-id="bf39e-653">- Content</span></span><br><span data-ttu-id="bf39e-654">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bf39e-654">
         - Taskpane</span></span><br><span data-ttu-id="bf39e-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bf39e-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="bf39e-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bf39e-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bf39e-658">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bf39e-658">-DocumentEvents</span></span><br><span data-ttu-id="bf39e-659">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-659">
         -HtmlCoercion</span></span><br><span data-ttu-id="bf39e-660">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-660">
         -ImageCoercion</span></span><br><span data-ttu-id="bf39e-661">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bf39e-661">
         - Settings</span></span><br><span data-ttu-id="bf39e-662">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bf39e-662">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="bf39e-663">См. также</span><span class="sxs-lookup"><span data-stu-id="bf39e-663">See also</span></span>

- [<span data-ttu-id="bf39e-664">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bf39e-664">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="bf39e-665">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bf39e-665">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="bf39e-666">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="bf39e-666">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="bf39e-667">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="bf39e-667">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
