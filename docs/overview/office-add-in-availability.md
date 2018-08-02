---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы требований для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 07/31/2018
ms.openlocfilehash: 084029c0a5b70b73eaa0b3fcc180f4a813fb8b72
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703912"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="aed92-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="aed92-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="aed92-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="aed92-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="aed92-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="aed92-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="aed92-106">Символ \* (звездочка) в ячейке таблицы указывает, что поддержка скоро появится.</span><span class="sxs-lookup"><span data-stu-id="aed92-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="aed92-107">С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="aed92-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="aed92-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="aed92-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="aed92-110">Excel</span><span class="sxs-lookup"><span data-stu-id="aed92-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="aed92-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="aed92-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="aed92-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="aed92-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="aed92-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="aed92-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="aed92-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="aed92-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="aed92-115">Office Online</span></span></td>
    <td> <span data-ttu-id="aed92-116">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-116">- Taskpane</span></span><br><span data-ttu-id="aed92-117">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-117">
        - Content</span></span><br><span data-ttu-id="aed92-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="aed92-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="aed92-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aed92-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aed92-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aed92-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aed92-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aed92-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aed92-125">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aed92-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="aed92-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="aed92-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-127">
        -BindingEvents</span></span><br><span data-ttu-id="aed92-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-128">
        -DocumentEvents</span></span><br><span data-ttu-id="aed92-129">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-129">
        -MatrixBindings</span></span><br><span data-ttu-id="aed92-130">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-130">
        -MatrixCoercion</span></span><br><span data-ttu-id="aed92-131">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-131">
        -TableBindings</span></span><br><span data-ttu-id="aed92-132">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-132">
        -TableCoercion</span></span><br><span data-ttu-id="aed92-133">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-133">
        -TextBindings</span></span><br><span data-ttu-id="aed92-134">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-134">
        -CompressedFile</span></span><br><span data-ttu-id="aed92-135">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-135">
        - Settings</span></span><br><span data-ttu-id="aed92-136">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-136">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-137">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-137">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="aed92-138">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-138">
        - Taskpane</span></span><br><span data-ttu-id="aed92-139">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-139">
        - Content</span></span></td>
    <td>  <span data-ttu-id="aed92-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="aed92-141">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-141">
        -BindingEvents</span></span><br><span data-ttu-id="aed92-142">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-142">
        -DocumentEvents</span></span><br><span data-ttu-id="aed92-143">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-143">
        -MatrixBindings</span></span><br><span data-ttu-id="aed92-144">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-144">
        -MatrixCoercion</span></span><br><span data-ttu-id="aed92-145">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-145">
        -TableBindings</span></span><br><span data-ttu-id="aed92-146">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-146">
        -TableCoercion</span></span><br><span data-ttu-id="aed92-147">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-147">
        -TextBindings</span></span><br><span data-ttu-id="aed92-148">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-148">
        - Settings</span></span><br><span data-ttu-id="aed92-149">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-149">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-150">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-150">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="aed92-151">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-151">- Taskpane</span></span><br><span data-ttu-id="aed92-152">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-152">
        - Content</span></span><br><span data-ttu-id="aed92-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="aed92-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aed92-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aed92-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aed92-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aed92-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aed92-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aed92-160">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aed92-160">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="aed92-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="aed92-162">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-162">-BindingEvents</span></span><br><span data-ttu-id="aed92-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-163">
        -DocumentEvents</span></span><br><span data-ttu-id="aed92-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-164">
        -MatrixBindings</span></span><br><span data-ttu-id="aed92-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-165">
        -MatrixCoercion</span></span><br><span data-ttu-id="aed92-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-166">
        -TableBindings</span></span><br><span data-ttu-id="aed92-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-167">
        -TableCoercion</span></span><br><span data-ttu-id="aed92-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-168">
        -TextBindings</span></span><br><span data-ttu-id="aed92-169">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-169">
        - Settings</span></span><br><span data-ttu-id="aed92-170">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-170">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-171">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="aed92-171">Office for iOS</span></span></td>
    <td><span data-ttu-id="aed92-172">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-172">- Taskpane</span></span><br><span data-ttu-id="aed92-173">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-173">
        - Content</span></span></td>
    <td><span data-ttu-id="aed92-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aed92-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aed92-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aed92-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aed92-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aed92-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aed92-180">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aed92-180">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="aed92-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="aed92-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-182">-BindingEvents</span></span><br><span data-ttu-id="aed92-183">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-183">
        -DocumentEvents</span></span><br><span data-ttu-id="aed92-184">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-184">
        -MatrixBindings</span></span><br><span data-ttu-id="aed92-185">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-185">
        -MatrixCoercion</span></span><br><span data-ttu-id="aed92-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-186">
        -TableBindings</span></span><br><span data-ttu-id="aed92-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-187">
        -TableCoercion</span></span><br><span data-ttu-id="aed92-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-188">
        -TextBindings</span></span><br><span data-ttu-id="aed92-189">
        - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-189">
        - Settings</span></span><br><span data-ttu-id="aed92-190">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-190">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-191">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="aed92-191">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="aed92-192">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-192">- Taskpane</span></span><br><span data-ttu-id="aed92-193">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-193">
        - Content</span></span><br><span data-ttu-id="aed92-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="aed92-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="aed92-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="aed92-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="aed92-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="aed92-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="aed92-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="aed92-201">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="aed92-201">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="aed92-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="aed92-203">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-203">-BindingEvents</span></span><br><span data-ttu-id="aed92-204">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-204">
        -DocumentEvents</span></span><br><span data-ttu-id="aed92-205">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-205">
        -MatrixBindings</span></span><br><span data-ttu-id="aed92-206">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-206">
        -MatrixCoercion</span></span><br><span data-ttu-id="aed92-207">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-207">
        -TableBindings</span></span><br><span data-ttu-id="aed92-208">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-208">
        -TableCoercion</span></span><br><span data-ttu-id="aed92-209">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-209">
        -TextBindings</span></span><br><span data-ttu-id="aed92-210">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-210">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="aed92-211">Outlook</span><span class="sxs-lookup"><span data-stu-id="aed92-211">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aed92-212">Платформа</span><span class="sxs-lookup"><span data-stu-id="aed92-212">Platform</span></span></th>
    <th><span data-ttu-id="aed92-213">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="aed92-213">Extension points</span></span></th> 
    <th><span data-ttu-id="aed92-214">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="aed92-214">API requirement sets</span></span></th> 
    <th><span data-ttu-id="aed92-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="aed92-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-216">Office Online</span><span class="sxs-lookup"><span data-stu-id="aed92-216">Office Online</span></span></td>
    <td> <span data-ttu-id="aed92-217">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="aed92-217">- Mail Read</span></span><br><span data-ttu-id="aed92-218">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="aed92-218">
      - Mail Compose</span></span><br><span data-ttu-id="aed92-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aed92-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aed92-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aed92-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aed92-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aed92-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="aed92-226">Недоступно</span><span class="sxs-lookup"><span data-stu-id="aed92-226">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-227">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-227">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="aed92-228">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="aed92-228">- Mail Read</span></span><br><span data-ttu-id="aed92-229">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="aed92-229">
      - Mail Compose</span></span><br><span data-ttu-id="aed92-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aed92-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aed92-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aed92-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="aed92-235">Недоступно</span><span class="sxs-lookup"><span data-stu-id="aed92-235">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-236">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-236">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="aed92-237">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="aed92-237">- Mail Read</span></span><br><span data-ttu-id="aed92-238">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="aed92-238">
      - Mail Compose</span></span><br><span data-ttu-id="aed92-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="aed92-240">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="aed92-240">
      - Modules</span></span></td>
    <td> <span data-ttu-id="aed92-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aed92-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aed92-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aed92-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aed92-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aed92-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="aed92-247">Недоступно</span><span class="sxs-lookup"><span data-stu-id="aed92-247">Not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-248">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="aed92-248">Office for iOS</span></span></td>
    <td> <span data-ttu-id="aed92-249">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="aed92-249">- Mail Read</span></span><br><span data-ttu-id="aed92-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aed92-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aed92-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aed92-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aed92-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="aed92-256">Недоступно</span><span class="sxs-lookup"><span data-stu-id="aed92-256">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-257">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="aed92-257">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="aed92-258">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="aed92-258">- Mail Read</span></span><br><span data-ttu-id="aed92-259">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="aed92-259">
      - Mail Compose</span></span><br><span data-ttu-id="aed92-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aed92-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aed92-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aed92-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aed92-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="aed92-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="aed92-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="aed92-267">Недоступно</span><span class="sxs-lookup"><span data-stu-id="aed92-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-268">Office для Android</span><span class="sxs-lookup"><span data-stu-id="aed92-268">Office for Android</span></span></td>
    <td> <span data-ttu-id="aed92-269">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="aed92-269">- Mail Read</span></span><br><span data-ttu-id="aed92-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="aed92-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="aed92-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="aed92-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="aed92-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="aed92-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="aed92-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="aed92-276">Недоступно</span><span class="sxs-lookup"><span data-stu-id="aed92-276">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="aed92-277">Слово</span><span class="sxs-lookup"><span data-stu-id="aed92-277">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aed92-278">Платформа</span><span class="sxs-lookup"><span data-stu-id="aed92-278">Platform</span></span></th>
    <th><span data-ttu-id="aed92-279">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="aed92-279">Extension points</span></span></th> 
    <th><span data-ttu-id="aed92-280">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="aed92-280">API requirement sets</span></span></th> 
    <th><span data-ttu-id="aed92-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="aed92-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-282">Office Online</span><span class="sxs-lookup"><span data-stu-id="aed92-282">Office Online</span></span></td>
    <td> <span data-ttu-id="aed92-283">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-283">- Taskpane</span></span><br><span data-ttu-id="aed92-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="aed92-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="aed92-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="aed92-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-289">-BindingEvents</span></span><br><span data-ttu-id="aed92-290">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="aed92-290">customXmlParts</span></span><br><span data-ttu-id="aed92-291">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-291">
         -MatrixBindings</span></span><br><span data-ttu-id="aed92-292">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-292">
         -MatrixCoercion</span></span><br><span data-ttu-id="aed92-293">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-293">
         -TableBindings</span></span><br><span data-ttu-id="aed92-294">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-294">
         -TableCoercion</span></span><br><span data-ttu-id="aed92-295">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-295">
         -TextBindings</span></span><br><span data-ttu-id="aed92-296">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-296">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-297">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aed92-297">
         -TextFile</span></span><br><span data-ttu-id="aed92-298">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-298">
         -ImageCoercion</span></span><br><span data-ttu-id="aed92-299">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-299">
         - Settings</span></span><br><span data-ttu-id="aed92-300">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-300">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-301">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-301">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="aed92-302">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-302">- Taskpane</span></span></td>
    <td> <span data-ttu-id="aed92-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-304">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-304">-BindingEvents</span></span><br><span data-ttu-id="aed92-305">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-305">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-306">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="aed92-306">
         -CustomXmlPart</span></span><br><span data-ttu-id="aed92-307">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-307">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-308">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-308">
         - File</span></span><br><span data-ttu-id="aed92-309">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-309">
         -HtmlCoercion</span></span><br><span data-ttu-id="aed92-310">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-310">
         -ImageCoercion</span></span><br><span data-ttu-id="aed92-311">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-311">
         -OoxmlCoercion</span></span><br><span data-ttu-id="aed92-312">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-312">
         -TableBindings</span></span><br><span data-ttu-id="aed92-313">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-313">
         -TableCoercion</span></span><br><span data-ttu-id="aed92-314">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-314">
         -TextBindings</span></span><br><span data-ttu-id="aed92-315">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aed92-315">
         -TextFile</span></span><br><span data-ttu-id="aed92-316">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-316">
         - Settings</span></span><br><span data-ttu-id="aed92-317">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-317">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-318">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-318">
         -MatrixCoercion</span></span><br><span data-ttu-id="aed92-319">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="aed92-319">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-320">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-320">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="aed92-321">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-321">- Taskpane</span></span><br><span data-ttu-id="aed92-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="aed92-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="aed92-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="aed92-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-327">-BindingEvents</span></span><br><span data-ttu-id="aed92-328">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-328">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-329">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="aed92-329">
         -CustomXmlPart</span></span><br><span data-ttu-id="aed92-330">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-330">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-331">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-331">
         - File</span></span><br><span data-ttu-id="aed92-332">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-332">
         -HtmlCoercion</span></span><br><span data-ttu-id="aed92-333">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-333">
         -ImageCoercion</span></span><br><span data-ttu-id="aed92-334">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-334">
         -OoxmlCoercion</span></span><br><span data-ttu-id="aed92-335">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-335">
         -TableBindings</span></span><br><span data-ttu-id="aed92-336">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-336">
         -TableCoercion</span></span><br><span data-ttu-id="aed92-337">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-337">
         -TextBindings</span></span><br><span data-ttu-id="aed92-338">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aed92-338">
         -TextFile</span></span><br><span data-ttu-id="aed92-339">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-339">
         - Settings</span></span><br><span data-ttu-id="aed92-340">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-340">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-341">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-341">
         -MatrixCoercion</span></span><br><span data-ttu-id="aed92-342">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="aed92-342">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-343">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="aed92-343">Office for iOS</span></span></td>
    <td> <span data-ttu-id="aed92-344">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-344">- Taskpane</span></span></td>
    <td> <span data-ttu-id="aed92-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="aed92-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="aed92-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="aed92-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="aed92-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="aed92-349">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-349">-BindingEvents</span></span><br><span data-ttu-id="aed92-350">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-350">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-351">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="aed92-351">
         -CustomXmlPart</span></span><br><span data-ttu-id="aed92-352">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-352">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-353">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-353">
         - File</span></span><br><span data-ttu-id="aed92-354">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-354">
         -HtmlCoercion</span></span><br><span data-ttu-id="aed92-355">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-355">
         -ImageCoercion</span></span><br><span data-ttu-id="aed92-356">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-356">
         -OoxmlCoercion</span></span><br><span data-ttu-id="aed92-357">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-357">
         -TableBindings</span></span><br><span data-ttu-id="aed92-358">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-358">
         -TableCoercion</span></span><br><span data-ttu-id="aed92-359">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-359">
         -TextBindings</span></span><br><span data-ttu-id="aed92-360">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aed92-360">
         -TextFile</span></span><br><span data-ttu-id="aed92-361">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-361">
         - Settings</span></span><br><span data-ttu-id="aed92-362">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-362">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="aed92-364">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="aed92-364">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-365">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="aed92-365">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="aed92-366">- Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-366">- Taskpane</span></span><br><span data-ttu-id="aed92-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="aed92-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="aed92-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="aed92-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="aed92-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="aed92-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="aed92-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="aed92-372">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-372">-BindingEvents</span></span><br><span data-ttu-id="aed92-373">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-373">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-374">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="aed92-374">
         -CustomXmlPart</span></span><br><span data-ttu-id="aed92-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-375">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-376">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-376">
         - File</span></span><br><span data-ttu-id="aed92-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-377">
         -HtmlCoercion</span></span><br><span data-ttu-id="aed92-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-378">
         -ImageCoercion</span></span><br><span data-ttu-id="aed92-379">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-379">
         -OoxmlCoercion</span></span><br><span data-ttu-id="aed92-380">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-380">
         -TableBindings</span></span><br><span data-ttu-id="aed92-381">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-381">
         -TableCoercion</span></span><br><span data-ttu-id="aed92-382">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="aed92-382">
         -TextBindings</span></span><br><span data-ttu-id="aed92-383">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="aed92-383">
         -TextFile</span></span><br><span data-ttu-id="aed92-384">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-384">
         - Settings</span></span><br><span data-ttu-id="aed92-385">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-385">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="aed92-387">
         - Привязки матрицы</span><span class="sxs-lookup"><span data-stu-id="aed92-387">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="aed92-388">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="aed92-388">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aed92-389">Платформа</span><span class="sxs-lookup"><span data-stu-id="aed92-389">Platform</span></span></th>
    <th><span data-ttu-id="aed92-390">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="aed92-390">Extension points</span></span></th> 
    <th><span data-ttu-id="aed92-391">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="aed92-391">API requirement sets</span></span></th> 
    <th><span data-ttu-id="aed92-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="aed92-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-393">Office Online</span><span class="sxs-lookup"><span data-stu-id="aed92-393">Office Online</span></span></td>
    <td> <span data-ttu-id="aed92-394">- Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-394">- Content</span></span><br><span data-ttu-id="aed92-395">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-395">
         - Taskpane</span></span><br><span data-ttu-id="aed92-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-398">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aed92-398">-ActiveView</span></span><br><span data-ttu-id="aed92-399">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-399">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-400">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-400">
         - File</span></span><br><span data-ttu-id="aed92-401">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="aed92-401">
         - Selection</span></span><br><span data-ttu-id="aed92-402">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-402">
         - Settings</span></span><br><span data-ttu-id="aed92-403">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-403">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-404">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-404">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-405">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-405">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="aed92-406">- Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-406">- Content</span></span><br><span data-ttu-id="aed92-407">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-407">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="aed92-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="aed92-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="aed92-409">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aed92-409">-ActiveView</span></span><br><span data-ttu-id="aed92-410">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-410">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-411">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-411">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-412">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-412">
         - File</span></span><br><span data-ttu-id="aed92-413">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="aed92-413">
         - Selection</span></span><br><span data-ttu-id="aed92-414">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-414">
         - Settings</span></span><br><span data-ttu-id="aed92-415">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-415">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-416">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="aed92-416">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="aed92-417">- Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-417">- Content</span></span><br><span data-ttu-id="aed92-418">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-418">
         - Taskpane</span></span><br><span data-ttu-id="aed92-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-421">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aed92-421">-ActiveView</span></span><br><span data-ttu-id="aed92-422">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-422">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-423">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-423">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-424">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-424">
         - File</span></span><br><span data-ttu-id="aed92-425">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="aed92-425">
         - Selection</span></span><br><span data-ttu-id="aed92-426">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-426">
         - Settings</span></span><br><span data-ttu-id="aed92-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-427">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-428">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-428">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-429">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="aed92-429">Office for iOS</span></span></td>
    <td> <span data-ttu-id="aed92-430">- Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-430">- Content</span></span><br><span data-ttu-id="aed92-431">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-431">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="aed92-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="aed92-433">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aed92-433">-ActiveView</span></span><br><span data-ttu-id="aed92-434">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-434">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-435">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-435">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-436">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-436">
         - File</span></span><br><span data-ttu-id="aed92-437">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="aed92-437">
         - Selection</span></span><br><span data-ttu-id="aed92-438">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-438">
         - Settings</span></span><br><span data-ttu-id="aed92-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-439">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-440">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-440">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-441">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="aed92-441">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="aed92-442">- Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-442">- Content</span></span><br><span data-ttu-id="aed92-443">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-443">
         - Taskpane</span></span><br><span data-ttu-id="aed92-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-446">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="aed92-446">-ActiveView</span></span><br><span data-ttu-id="aed92-447">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="aed92-447">
         -CompressedFile</span></span><br><span data-ttu-id="aed92-448">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-448">
         -DocumentEvents</span></span><br><span data-ttu-id="aed92-449">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="aed92-449">
         - File</span></span><br><span data-ttu-id="aed92-450">
         - Выделение</span><span class="sxs-lookup"><span data-stu-id="aed92-450">
         - Selection</span></span><br><span data-ttu-id="aed92-451">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-451">
         - Settings</span></span><br><span data-ttu-id="aed92-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-452">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-453">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="aed92-454">OneNote</span><span class="sxs-lookup"><span data-stu-id="aed92-454">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="aed92-455">Платформа</span><span class="sxs-lookup"><span data-stu-id="aed92-455">Platform</span></span></th>
    <th><span data-ttu-id="aed92-456">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="aed92-456">Extension points</span></span></th> 
    <th><span data-ttu-id="aed92-457">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="aed92-457">API requirement sets</span></span></th> 
    <th><span data-ttu-id="aed92-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="aed92-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="aed92-459">Office Online</span><span class="sxs-lookup"><span data-stu-id="aed92-459">Office Online</span></span></td>
    <td> <span data-ttu-id="aed92-460">- Контент</span><span class="sxs-lookup"><span data-stu-id="aed92-460">- Content</span></span><br><span data-ttu-id="aed92-461">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="aed92-461">
         - Taskpane</span></span><br><span data-ttu-id="aed92-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="aed92-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="aed92-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="aed92-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="aed92-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="aed92-465">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="aed92-465">-DocumentEvents</span></span><br><span data-ttu-id="aed92-466">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="aed92-466">
         - Settings</span></span><br><span data-ttu-id="aed92-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-467">
         -TextCoercion</span></span><br><span data-ttu-id="aed92-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="aed92-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="aed92-469">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="aed92-470">См. также</span><span class="sxs-lookup"><span data-stu-id="aed92-470">See also</span></span>

- [<span data-ttu-id="aed92-471">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="aed92-471">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="aed92-472">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="aed92-472">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="aed92-473">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="aed92-473">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="aed92-474">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="aed92-474">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

