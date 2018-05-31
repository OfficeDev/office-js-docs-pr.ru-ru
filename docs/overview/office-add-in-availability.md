---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы требований для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438874"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c3d6e-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c3d6e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c3d6e-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="c3d6e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="c3d6e-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="c3d6e-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="c3d6e-106">Символ \* (звездочка) в ячейке таблицы указывает, что поддержка скоро появится.</span><span class="sxs-lookup"><span data-stu-id="c3d6e-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="c3d6e-107">С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="c3d6e-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="c3d6e-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="c3d6e-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="c3d6e-110">Excel</span><span class="sxs-lookup"><span data-stu-id="c3d6e-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c3d6e-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="c3d6e-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c3d6e-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c3d6e-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="c3d6e-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c3d6e-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="c3d6e-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="c3d6e-115">Office Online</span></span></td>
    <td> <span data-ttu-id="c3d6e-116">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-116">- Taskpane</span></span><br><span data-ttu-id="c3d6e-117">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-117">
        - Content</span></span><br><span data-ttu-id="c3d6e-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="c3d6e-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c3d6e-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d6e-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d6e-124">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-124">
        -BindingEvents</span></span><br><span data-ttu-id="c3d6e-125">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-125">
        -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-126">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-126">
        -MatrixBindings</span></span><br><span data-ttu-id="c3d6e-127">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-127">
        -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-128">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-128">
        -TableBindings</span></span><br><span data-ttu-id="c3d6e-129">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-129">
        -TableCoercion</span></span><br><span data-ttu-id="c3d6e-130">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-130">
        -TextBindings</span></span><br><span data-ttu-id="c3d6e-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-131">
        -CompressedFile</span></span><br><span data-ttu-id="c3d6e-132">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-132">
        - Settings</span></span><br><span data-ttu-id="c3d6e-133">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-133">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-134">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-134">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="c3d6e-135">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-135">
        - Taskpane</span></span><br><span data-ttu-id="c3d6e-136">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-136">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c3d6e-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d6e-138">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-138">
        -BindingEvents</span></span><br><span data-ttu-id="c3d6e-139">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-139">
        -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-140">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-140">
        -MatrixBindings</span></span><br><span data-ttu-id="c3d6e-141">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-141">
        -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-142">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-142">
        -TableBindings</span></span><br><span data-ttu-id="c3d6e-143">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-143">
        -TableCoercion</span></span><br><span data-ttu-id="c3d6e-144">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-144">
        -TextBindings</span></span><br><span data-ttu-id="c3d6e-145">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-145">
        - Settings</span></span><br><span data-ttu-id="c3d6e-146">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-146">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-147">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-147">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="c3d6e-148">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-148">- Taskpane</span></span><br><span data-ttu-id="c3d6e-149">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-149">
        - Content</span></span><br><span data-ttu-id="c3d6e-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c3d6e-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c3d6e-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d6e-156">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-156">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-157">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-157">
        -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-158">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-158">
        -MatrixBindings</span></span><br><span data-ttu-id="c3d6e-159">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-159">
        -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-160">
        -TableBindings</span></span><br><span data-ttu-id="c3d6e-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-161">
        -TableCoercion</span></span><br><span data-ttu-id="c3d6e-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-162">
        -TextBindings</span></span><br><span data-ttu-id="c3d6e-163">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-163">
        - Settings</span></span><br><span data-ttu-id="c3d6e-164">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-164">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-165">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="c3d6e-165">Office for iOS</span></span></td>
    <td><span data-ttu-id="c3d6e-166">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-166">- Taskpane</span></span><br><span data-ttu-id="c3d6e-167">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-167">
        - Content</span></span></td>
    <td><span data-ttu-id="c3d6e-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d6e-172">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-172">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-173">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-173">
        -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-174">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-174">
        -MatrixBindings</span></span><br><span data-ttu-id="c3d6e-175">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-175">
        -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-176">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-176">
        -TableBindings</span></span><br><span data-ttu-id="c3d6e-177">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-177">
        -TableCoercion</span></span><br><span data-ttu-id="c3d6e-178">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-178">
        -TextBindings</span></span><br><span data-ttu-id="c3d6e-179">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-179">
        - Settings</span></span><br><span data-ttu-id="c3d6e-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-181">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c3d6e-181">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="c3d6e-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-182">- Taskpane</span></span><br><span data-ttu-id="c3d6e-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-183">
        - Content</span></span><br><span data-ttu-id="c3d6e-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c3d6e-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c3d6e-189">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-189">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-190">
        -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-191">
        -MatrixBindings</span></span><br><span data-ttu-id="c3d6e-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-192">
        -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-193">
        -TableBindings</span></span><br><span data-ttu-id="c3d6e-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-194">
        -TableCoercion</span></span><br><span data-ttu-id="c3d6e-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-195">
        -TextBindings</span></span><br><span data-ttu-id="c3d6e-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-196">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="c3d6e-197">Outlook</span><span class="sxs-lookup"><span data-stu-id="c3d6e-197">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d6e-198">Платформа</span><span class="sxs-lookup"><span data-stu-id="c3d6e-198">Platform</span></span></th>
    <th><span data-ttu-id="c3d6e-199">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c3d6e-199">Extension points</span></span></th> 
    <th><span data-ttu-id="c3d6e-200">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c3d6e-200">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c3d6e-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-202">Office Online</span><span class="sxs-lookup"><span data-stu-id="c3d6e-202">Office Online</span></span></td>
    <td> <span data-ttu-id="c3d6e-203">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-203">- Mail Read</span></span><br><span data-ttu-id="c3d6e-204">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-204">
      - Mail Compose</span></span><br><span data-ttu-id="c3d6e-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d6e-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d6e-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d6e-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d6e-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d6e-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c3d6e-212">Недоступен</span><span class="sxs-lookup"><span data-stu-id="c3d6e-212">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-213">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-213">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c3d6e-214">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-214">- Mail Read</span></span><br><span data-ttu-id="c3d6e-215">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-215">
      - Mail Compose</span></span><br><span data-ttu-id="c3d6e-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d6e-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d6e-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d6e-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="c3d6e-221">Недоступен</span><span class="sxs-lookup"><span data-stu-id="c3d6e-221">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-222">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-222">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c3d6e-223">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-223">- Mail Read</span></span><br><span data-ttu-id="c3d6e-224">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-224">
      - Mail Compose</span></span><br><span data-ttu-id="c3d6e-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c3d6e-226">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="c3d6e-226">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c3d6e-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d6e-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d6e-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d6e-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d6e-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d6e-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c3d6e-233">Недоступен</span><span class="sxs-lookup"><span data-stu-id="c3d6e-233">not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-234">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="c3d6e-234">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c3d6e-235">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-235">- Mail Read</span></span><br><span data-ttu-id="c3d6e-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d6e-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d6e-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d6e-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d6e-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="c3d6e-242">Недоступен</span><span class="sxs-lookup"><span data-stu-id="c3d6e-242">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-243">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c3d6e-243">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c3d6e-244">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-244">- Mail Read</span></span><br><span data-ttu-id="c3d6e-245">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-245">
      - Mail Compose</span></span><br><span data-ttu-id="c3d6e-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d6e-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d6e-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d6e-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d6e-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c3d6e-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c3d6e-253">Недоступен</span><span class="sxs-lookup"><span data-stu-id="c3d6e-253">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-254">Office для Android</span><span class="sxs-lookup"><span data-stu-id="c3d6e-254">Office for Android</span></span></td>
    <td> <span data-ttu-id="c3d6e-255">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="c3d6e-255">- Mail Read</span></span><br><span data-ttu-id="c3d6e-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c3d6e-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c3d6e-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c3d6e-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c3d6e-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c3d6e-262">Недоступен</span><span class="sxs-lookup"><span data-stu-id="c3d6e-262">not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="c3d6e-263">Word</span><span class="sxs-lookup"><span data-stu-id="c3d6e-263">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d6e-264">Платформа</span><span class="sxs-lookup"><span data-stu-id="c3d6e-264">Platform</span></span></th>
    <th><span data-ttu-id="c3d6e-265">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c3d6e-265">Extension points</span></span></th> 
    <th><span data-ttu-id="c3d6e-266">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c3d6e-266">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c3d6e-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-268">Office Online</span><span class="sxs-lookup"><span data-stu-id="c3d6e-268">Office Online</span></span></td>
    <td> <span data-ttu-id="c3d6e-269">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-269">- Taskpane</span></span><br><span data-ttu-id="c3d6e-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-275">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-276">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c3d6e-276">customXmlParts</span></span><br><span data-ttu-id="c3d6e-277">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-277">
         -MatrixBindings</span></span><br><span data-ttu-id="c3d6e-278">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-278">
         -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-279">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-279">
         -TableBindings</span></span><br><span data-ttu-id="c3d6e-280">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-280">
         -TableCoercion</span></span><br><span data-ttu-id="c3d6e-281">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-281">
         -TextBindings</span></span><br><span data-ttu-id="c3d6e-282">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-282">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-283">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-283">
         -TextFile</span></span><br><span data-ttu-id="c3d6e-284">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-284">
         -ImageCoercion</span></span><br><span data-ttu-id="c3d6e-285">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-285">
         - Settings</span></span><br><span data-ttu-id="c3d6e-286">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-286">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-287">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c3d6e-288">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-288">- Taskpane</span></span></td>
    <td> <span data-ttu-id="c3d6e-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-290">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-291">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-291">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-292">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c3d6e-292">
         -CustomXmlPart</span></span><br><span data-ttu-id="c3d6e-293">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-293">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-294">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-294">
         - File</span></span><br><span data-ttu-id="c3d6e-295">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-295">
         -HtmlCoercion</span></span><br><span data-ttu-id="c3d6e-296">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-296">
         -ImageCoercion</span></span><br><span data-ttu-id="c3d6e-297">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-297">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c3d6e-298">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-298">
         -TableBindings</span></span><br><span data-ttu-id="c3d6e-299">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-299">
         -TableCoercion</span></span><br><span data-ttu-id="c3d6e-300">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-300">
         -TextBindings</span></span><br><span data-ttu-id="c3d6e-301">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-301">
         -TextFile</span></span><br><span data-ttu-id="c3d6e-302">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-302">
         - Settings</span></span><br><span data-ttu-id="c3d6e-303">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-303">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-304">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-304">
         -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-305">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="c3d6e-305">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-306">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-306">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c3d6e-307">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-307">- Taskpane</span></span><br><span data-ttu-id="c3d6e-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-313">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-313">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-314">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-314">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-315">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c3d6e-315">
         -CustomXmlPart</span></span><br><span data-ttu-id="c3d6e-316">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-316">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-317">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-317">
         - File</span></span><br><span data-ttu-id="c3d6e-318">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-318">
         -HtmlCoercion</span></span><br><span data-ttu-id="c3d6e-319">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-319">
         -ImageCoercion</span></span><br><span data-ttu-id="c3d6e-320">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-320">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c3d6e-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-321">
         -TableBindings</span></span><br><span data-ttu-id="c3d6e-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-322">
         -TableCoercion</span></span><br><span data-ttu-id="c3d6e-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-323">
         -TextBindings</span></span><br><span data-ttu-id="c3d6e-324">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-324">
         -TextFile</span></span><br><span data-ttu-id="c3d6e-325">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-325">
         - Settings</span></span><br><span data-ttu-id="c3d6e-326">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-326">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-327">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-327">
         -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-328">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="c3d6e-328">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-329">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="c3d6e-329">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c3d6e-330">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-330">- Taskpane</span></span></td>
    <td> <span data-ttu-id="c3d6e-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c3d6e-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c3d6e-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-335">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-336">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-336">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-337">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c3d6e-337">
         -CustomXmlPart</span></span><br><span data-ttu-id="c3d6e-338">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-338">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-339">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-339">
         - File</span></span><br><span data-ttu-id="c3d6e-340">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-340">
         -HtmlCoercion</span></span><br><span data-ttu-id="c3d6e-341">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-341">
         -ImageCoercion</span></span><br><span data-ttu-id="c3d6e-342">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-342">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c3d6e-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-343">
         -TableBindings</span></span><br><span data-ttu-id="c3d6e-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-344">
         -TableCoercion</span></span><br><span data-ttu-id="c3d6e-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-345">
         -TextBindings</span></span><br><span data-ttu-id="c3d6e-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-346">
         -TextFile</span></span><br><span data-ttu-id="c3d6e-347">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-347">
         - Settings</span></span><br><span data-ttu-id="c3d6e-348">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-348">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-349">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-349">
         -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-350">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="c3d6e-350">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-351">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c3d6e-351">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c3d6e-352">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-352">- Taskpane</span></span><br><span data-ttu-id="c3d6e-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c3d6e-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c3d6e-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c3d6e-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c3d6e-358">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-358">-BindingEvents</span></span><br><span data-ttu-id="c3d6e-359">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-359">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-360">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="c3d6e-360">
         -CustomXmlPart</span></span><br><span data-ttu-id="c3d6e-361">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-361">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-362">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-362">
         - File</span></span><br><span data-ttu-id="c3d6e-363">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-363">
         -HtmlCoercion</span></span><br><span data-ttu-id="c3d6e-364">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-364">
         -ImageCoercion</span></span><br><span data-ttu-id="c3d6e-365">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-365">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c3d6e-366">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-366">
         -TableBindings</span></span><br><span data-ttu-id="c3d6e-367">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-367">
         -TableCoercion</span></span><br><span data-ttu-id="c3d6e-368">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c3d6e-368">
         -TextBindings</span></span><br><span data-ttu-id="c3d6e-369">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-369">
         -TextFile</span></span><br><span data-ttu-id="c3d6e-370">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-370">
         - Settings</span></span><br><span data-ttu-id="c3d6e-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-371">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-372">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-372">
         -MatrixCoercion</span></span><br><span data-ttu-id="c3d6e-373">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="c3d6e-373">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c3d6e-374">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c3d6e-374">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d6e-375">Платформа</span><span class="sxs-lookup"><span data-stu-id="c3d6e-375">Platform</span></span></th>
    <th><span data-ttu-id="c3d6e-376">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c3d6e-376">Extension points</span></span></th> 
    <th><span data-ttu-id="c3d6e-377">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c3d6e-377">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c3d6e-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-379">Office Online</span><span class="sxs-lookup"><span data-stu-id="c3d6e-379">Office Online</span></span></td>
    <td> <span data-ttu-id="c3d6e-380">- Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-380">- Content</span></span><br><span data-ttu-id="c3d6e-381">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-381">
         - Taskpane</span></span><br><span data-ttu-id="c3d6e-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-384">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d6e-384">-ActiveView</span></span><br><span data-ttu-id="c3d6e-385">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-385">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-386">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-386">
         - File</span></span><br><span data-ttu-id="c3d6e-387">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="c3d6e-387">
         - Selection</span></span><br><span data-ttu-id="c3d6e-388">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-388">
         - Settings</span></span><br><span data-ttu-id="c3d6e-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-389">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-390">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-390">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-391">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c3d6e-392">- Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-392">- Content</span></span><br><span data-ttu-id="c3d6e-393">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-393">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="c3d6e-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c3d6e-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c3d6e-395">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d6e-395">-ActiveView</span></span><br><span data-ttu-id="c3d6e-396">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-396">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-397">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-398">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-398">
         - File</span></span><br><span data-ttu-id="c3d6e-399">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="c3d6e-399">
         - Selection</span></span><br><span data-ttu-id="c3d6e-400">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-400">
         - Settings</span></span><br><span data-ttu-id="c3d6e-401">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-401">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-402">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-402">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c3d6e-403">- Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-403">- Content</span></span><br><span data-ttu-id="c3d6e-404">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-404">
         - Taskpane</span></span><br><span data-ttu-id="c3d6e-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-407">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d6e-407">-ActiveView</span></span><br><span data-ttu-id="c3d6e-408">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-408">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-409">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-409">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-410">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-410">
         - File</span></span><br><span data-ttu-id="c3d6e-411">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="c3d6e-411">
         - Selection</span></span><br><span data-ttu-id="c3d6e-412">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-412">
         - Settings</span></span><br><span data-ttu-id="c3d6e-413">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-413">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-414">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-414">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-415">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="c3d6e-415">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c3d6e-416">- Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-416">- Content</span></span><br><span data-ttu-id="c3d6e-417">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-417">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="c3d6e-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="c3d6e-419">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d6e-419">-ActiveView</span></span><br><span data-ttu-id="c3d6e-420">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-420">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-421">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-421">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-422">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-422">
         - File</span></span><br><span data-ttu-id="c3d6e-423">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="c3d6e-423">
         - Selection</span></span><br><span data-ttu-id="c3d6e-424">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-424">
         - Settings</span></span><br><span data-ttu-id="c3d6e-425">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-425">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-426">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-426">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-427">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c3d6e-427">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c3d6e-428">- Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-428">- Content</span></span><br><span data-ttu-id="c3d6e-429">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-429">
         - Taskpane</span></span><br><span data-ttu-id="c3d6e-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c3d6e-432">-ActiveView</span></span><br><span data-ttu-id="c3d6e-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c3d6e-433">
         -CompressedFile</span></span><br><span data-ttu-id="c3d6e-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-434">
         -DocumentEvents</span></span><br><span data-ttu-id="c3d6e-435">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c3d6e-435">
         - File</span></span><br><span data-ttu-id="c3d6e-436">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="c3d6e-436">
         - Selection</span></span><br><span data-ttu-id="c3d6e-437">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-437">
         - Settings</span></span><br><span data-ttu-id="c3d6e-438">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-438">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-439">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-439">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="c3d6e-440">OneNote</span><span class="sxs-lookup"><span data-stu-id="c3d6e-440">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c3d6e-441">Платформа</span><span class="sxs-lookup"><span data-stu-id="c3d6e-441">Platform</span></span></th>
    <th><span data-ttu-id="c3d6e-442">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c3d6e-442">Extension points</span></span></th> 
    <th><span data-ttu-id="c3d6e-443">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c3d6e-443">API requirement sets</span></span></th> 
    <th><span data-ttu-id="c3d6e-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-445">Office Online</span><span class="sxs-lookup"><span data-stu-id="c3d6e-445">Office Online</span></span></td>
    <td> <span data-ttu-id="c3d6e-446">- Контент</span><span class="sxs-lookup"><span data-stu-id="c3d6e-446">- Content</span></span><br><span data-ttu-id="c3d6e-447">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c3d6e-447">
         - Taskpane</span></span><br><span data-ttu-id="c3d6e-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c3d6e-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c3d6e-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c3d6e-451">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c3d6e-451">-DocumentEvents</span></span><br><span data-ttu-id="c3d6e-452">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c3d6e-452">
         - Settings</span></span><br><span data-ttu-id="c3d6e-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-453">
         -TextCoercion</span></span><br><span data-ttu-id="c3d6e-454">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-454">
         -HtmlCoercion</span></span><br><span data-ttu-id="c3d6e-455">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c3d6e-455">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-456">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-456">Office 2013 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td><span data-ttu-id="c3d6e-457">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c3d6e-457">Office 2016 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-458">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="c3d6e-458">Office for iOS</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c3d6e-459">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c3d6e-459">Office 2016 for Mac</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

<span data-ttu-id="c3d6e-460">\* = поддержка скоро появится.</span><span class="sxs-lookup"><span data-stu-id="c3d6e-460">\* = We're working on it.</span></span> 

## <a name="see-also"></a><span data-ttu-id="c3d6e-461">См. также</span><span class="sxs-lookup"><span data-stu-id="c3d6e-461">See also</span></span>

- [<span data-ttu-id="c3d6e-462">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c3d6e-462">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c3d6e-463">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c3d6e-463">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="c3d6e-464">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="c3d6e-464">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="c3d6e-465">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="c3d6e-465">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

