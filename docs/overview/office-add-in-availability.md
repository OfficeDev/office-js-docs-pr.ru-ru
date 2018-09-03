---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы требований для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 08/30/2018
ms.openlocfilehash: 06fb073693bd8adca7d196f4361699ac3f54cee1
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797303"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="a1120-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a1120-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="a1120-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="a1120-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="a1120-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="a1120-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="a1120-106">Символ \* (звездочка) в ячейке таблицы указывает, что поддержка скоро появится.</span><span class="sxs-lookup"><span data-stu-id="a1120-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="a1120-107">С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="a1120-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="a1120-p103">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="a1120-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="a1120-110">Excel</span><span class="sxs-lookup"><span data-stu-id="a1120-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="a1120-111">Платформа</span><span class="sxs-lookup"><span data-stu-id="a1120-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="a1120-112">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="a1120-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="a1120-113">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="a1120-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="a1120-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="a1120-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="a1120-115">Office Online</span></span></td>
    <td> <span data-ttu-id="a1120-116">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-116">- Taskpane</span></span><br><span data-ttu-id="a1120-117">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-117">
        - Content</span></span><br><span data-ttu-id="a1120-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="a1120-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a1120-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a1120-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a1120-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a1120-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a1120-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a1120-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a1120-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a1120-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="a1120-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a1120-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-127">
        -BindingEvents</span></span><br><span data-ttu-id="a1120-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-128">
        -CompressedFile</span></span><br><span data-ttu-id="a1120-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-129">
        -DocumentEvents</span></span><br><span data-ttu-id="a1120-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="a1120-130">
        - File</span></span><br><span data-ttu-id="a1120-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-131">
        -MatrixBindings</span></span><br><span data-ttu-id="a1120-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="a1120-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-133">
        - Selection</span></span><br><span data-ttu-id="a1120-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-134">
        - Settings</span></span><br><span data-ttu-id="a1120-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-135">
        -TableBindings</span></span><br><span data-ttu-id="a1120-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-136">
        -TableCoercion</span></span><br><span data-ttu-id="a1120-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-137">
        -TextBindings</span></span><br><span data-ttu-id="a1120-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-139">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="a1120-140">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-140">
        - Taskpane</span></span><br><span data-ttu-id="a1120-141">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="a1120-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a1120-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-143">
        -BindingEvents</span></span><br><span data-ttu-id="a1120-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-144">
        -CompressedFile</span></span><br><span data-ttu-id="a1120-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-145">
        -DocumentEvents</span></span><br><span data-ttu-id="a1120-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="a1120-146">
        - File</span></span><br><span data-ttu-id="a1120-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-147">
        -ImageCoercion</span></span><br><span data-ttu-id="a1120-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-148">
        -MatrixBindings</span></span><br><span data-ttu-id="a1120-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="a1120-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-150">
        - Selection</span></span><br><span data-ttu-id="a1120-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-151">
        - Settings</span></span><br><span data-ttu-id="a1120-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-152">
        -TableBindings</span></span><br><span data-ttu-id="a1120-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-153">
        -TableCoercion</span></span><br><span data-ttu-id="a1120-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-154">
        -TextBindings</span></span><br><span data-ttu-id="a1120-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-156">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="a1120-157">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-157">- Taskpane</span></span><br><span data-ttu-id="a1120-158">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-158">
        - Content</span></span><br><span data-ttu-id="a1120-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a1120-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a1120-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a1120-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a1120-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a1120-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a1120-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a1120-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a1120-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="a1120-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a1120-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-168">-BindingEvents</span></span><br><span data-ttu-id="a1120-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-169">
        -CompressedFile</span></span><br><span data-ttu-id="a1120-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-170">
        -DocumentEvents</span></span><br><span data-ttu-id="a1120-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="a1120-171">
        - File</span></span><br><span data-ttu-id="a1120-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-172">
        -ImageCoercion</span></span><br><span data-ttu-id="a1120-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-173">
        -MatrixBindings</span></span><br><span data-ttu-id="a1120-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="a1120-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-175">
        - Selection</span></span><br><span data-ttu-id="a1120-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-176">
        - Settings</span></span><br><span data-ttu-id="a1120-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-177">
        -TableBindings</span></span><br><span data-ttu-id="a1120-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-178">
        -TableCoercion</span></span><br><span data-ttu-id="a1120-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-179">
        -TextBindings</span></span><br><span data-ttu-id="a1120-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-181">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="a1120-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="a1120-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-182">- Taskpane</span></span><br><span data-ttu-id="a1120-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-183">
        - Content</span></span></td>
    <td><span data-ttu-id="a1120-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a1120-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a1120-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a1120-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a1120-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a1120-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a1120-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a1120-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="a1120-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a1120-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-192">-BindingEvents</span></span><br><span data-ttu-id="a1120-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-193">
        -CompressedFile</span></span><br><span data-ttu-id="a1120-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-194">
        -DocumentEvents</span></span><br><span data-ttu-id="a1120-195">
        - File</span><span class="sxs-lookup"><span data-stu-id="a1120-195">
        - File</span></span><br><span data-ttu-id="a1120-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-196">
        -ImageCoercion</span></span><br><span data-ttu-id="a1120-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-197">
        -MatrixBindings</span></span><br><span data-ttu-id="a1120-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="a1120-199">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-199">
        - Selection</span></span><br><span data-ttu-id="a1120-200">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-200">
        - Settings</span></span><br><span data-ttu-id="a1120-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-201">
        -TableBindings</span></span><br><span data-ttu-id="a1120-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-202">
        -TableCoercion</span></span><br><span data-ttu-id="a1120-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-203">
        -TextBindings</span></span><br><span data-ttu-id="a1120-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-205">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="a1120-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="a1120-206">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-206">- Taskpane</span></span><br><span data-ttu-id="a1120-207">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-207">
        - Content</span></span><br><span data-ttu-id="a1120-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a1120-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a1120-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a1120-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a1120-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a1120-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a1120-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a1120-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a1120-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="a1120-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a1120-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-217">-BindingEvents</span></span><br><span data-ttu-id="a1120-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-218">
        -CompressedFile</span></span><br><span data-ttu-id="a1120-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-219">
        -DocumentEvents</span></span><br><span data-ttu-id="a1120-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="a1120-220">
        - File</span></span><br><span data-ttu-id="a1120-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-221">
        -ImageCoercion</span></span><br><span data-ttu-id="a1120-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-222">
        -MatrixBindings</span></span><br><span data-ttu-id="a1120-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="a1120-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-224">
        -PdfFile</span></span><br><span data-ttu-id="a1120-225">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-225">
        - Selection</span></span><br><span data-ttu-id="a1120-226">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-226">
        - Settings</span></span><br><span data-ttu-id="a1120-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-227">
        -TableBindings</span></span><br><span data-ttu-id="a1120-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-228">
        -TableCoercion</span></span><br><span data-ttu-id="a1120-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-229">
        -TextBindings</span></span><br><span data-ttu-id="a1120-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="a1120-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="a1120-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a1120-232">Платформа</span><span class="sxs-lookup"><span data-stu-id="a1120-232">Platform</span></span></th>
    <th><span data-ttu-id="a1120-233">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="a1120-233">Extension points</span></span></th>
    <th><span data-ttu-id="a1120-234">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="a1120-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="a1120-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="a1120-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="a1120-236">Office Online</span></span></td>
    <td> <span data-ttu-id="a1120-237">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="a1120-237">- Mail Read</span></span><br><span data-ttu-id="a1120-238">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="a1120-238">
      - Mail Compose</span></span><br><span data-ttu-id="a1120-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a1120-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a1120-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a1120-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a1120-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a1120-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a1120-246">Недоступна</span><span class="sxs-lookup"><span data-stu-id="a1120-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-247">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a1120-248">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="a1120-248">- Mail Read</span></span><br><span data-ttu-id="a1120-249">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="a1120-249">
      - Mail Compose</span></span><br><span data-ttu-id="a1120-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a1120-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a1120-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a1120-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="a1120-255">Недоступна</span><span class="sxs-lookup"><span data-stu-id="a1120-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-256">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a1120-257">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="a1120-257">- Mail Read</span></span><br><span data-ttu-id="a1120-258">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="a1120-258">
      - Mail Compose</span></span><br><span data-ttu-id="a1120-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a1120-260">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="a1120-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a1120-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a1120-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a1120-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a1120-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a1120-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a1120-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a1120-267">Недоступна</span><span class="sxs-lookup"><span data-stu-id="a1120-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-268">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="a1120-268">Office for iOS</span></span></td>
    <td> <span data-ttu-id="a1120-269">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="a1120-269">- Mail Read</span></span><br><span data-ttu-id="a1120-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a1120-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a1120-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a1120-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a1120-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a1120-276">Недоступна</span><span class="sxs-lookup"><span data-stu-id="a1120-276">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-277">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="a1120-277">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="a1120-278">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="a1120-278">- Mail Read</span></span><br><span data-ttu-id="a1120-279">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="a1120-279">
      - Mail Compose</span></span><br><span data-ttu-id="a1120-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a1120-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a1120-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a1120-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a1120-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a1120-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a1120-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a1120-287">Недоступна</span><span class="sxs-lookup"><span data-stu-id="a1120-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-288">Office для Android</span><span class="sxs-lookup"><span data-stu-id="a1120-288">Office for Android</span></span></td>
    <td> <span data-ttu-id="a1120-289">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="a1120-289">- Mail Read</span></span><br><span data-ttu-id="a1120-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a1120-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a1120-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a1120-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a1120-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a1120-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a1120-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a1120-296">Недоступна</span><span class="sxs-lookup"><span data-stu-id="a1120-296">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="a1120-297">Word</span><span class="sxs-lookup"><span data-stu-id="a1120-297">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a1120-298">Платформа</span><span class="sxs-lookup"><span data-stu-id="a1120-298">Platform</span></span></th>
    <th><span data-ttu-id="a1120-299">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="a1120-299">Extension points</span></span></th>
    <th><span data-ttu-id="a1120-300">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="a1120-300">API requirement sets</span></span></th>
    <th><span data-ttu-id="a1120-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="a1120-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-302">Office Online</span><span class="sxs-lookup"><span data-stu-id="a1120-302">Office Online</span></span></td>
    <td> <span data-ttu-id="a1120-303">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-303">- Taskpane</span></span><br><span data-ttu-id="a1120-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a1120-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a1120-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a1120-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-309">-BindingEvents</span></span><br><span data-ttu-id="a1120-310">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a1120-310">customXmlParts</span></span><br><span data-ttu-id="a1120-311">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-311">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-312">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-312">
         - File</span></span><br><span data-ttu-id="a1120-313">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-313">
         -HtmlCoercion</span></span><br><span data-ttu-id="a1120-314">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-314">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-315">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-315">
         -MatrixBindings</span></span><br><span data-ttu-id="a1120-316">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-316">
         -MatrixCoercion</span></span><br><span data-ttu-id="a1120-317">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-317">
         -OoxmlCoercion</span></span><br><span data-ttu-id="a1120-318">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-318">
         -PdfFile</span></span><br><span data-ttu-id="a1120-319">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-319">
         - Selection</span></span><br><span data-ttu-id="a1120-320">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-320">
         - Settings</span></span><br><span data-ttu-id="a1120-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-321">
         -TableBindings</span></span><br><span data-ttu-id="a1120-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-322">
         -TableCoercion</span></span><br><span data-ttu-id="a1120-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-323">
         -TextBindings</span></span><br><span data-ttu-id="a1120-324">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-324">
         -TextCoercion</span></span><br><span data-ttu-id="a1120-325">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a1120-325">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-326">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-326">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a1120-327">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-327">- Taskpane</span></span></td>
    <td> <span data-ttu-id="a1120-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-329">-BindingEvents</span></span><br><span data-ttu-id="a1120-330">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-330">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-331">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a1120-331">customXmlParts</span></span><br><span data-ttu-id="a1120-332">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-332">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-333">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-333">
         - File</span></span><br><span data-ttu-id="a1120-334">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-334">
         -HtmlCoercion</span></span><br><span data-ttu-id="a1120-335">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-335">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-336">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-336">
         -MatrixBindings</span></span><br><span data-ttu-id="a1120-337">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-337">
         -MatrixCoercion</span></span><br><span data-ttu-id="a1120-338">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-338">
         -OoxmlCoercion</span></span><br><span data-ttu-id="a1120-339">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-339">
         -PdfFile</span></span><br><span data-ttu-id="a1120-340">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-340">
         - Selection</span></span><br><span data-ttu-id="a1120-341">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-341">
         - Settings</span></span><br><span data-ttu-id="a1120-342">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-342">
         -TableBindings</span></span><br><span data-ttu-id="a1120-343">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-343">
         -TableCoercion</span></span><br><span data-ttu-id="a1120-344">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-344">
         -TextBindings</span></span><br><span data-ttu-id="a1120-345">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-345">
         -TextCoercion</span></span><br><span data-ttu-id="a1120-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a1120-346">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-347">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-347">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a1120-348">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-348">- Taskpane</span></span><br><span data-ttu-id="a1120-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a1120-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a1120-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a1120-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-354">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-354">-BindingEvents</span></span><br><span data-ttu-id="a1120-355">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-355">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-356">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a1120-356">customXmlParts</span></span><br><span data-ttu-id="a1120-357">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-357">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-358">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-358">
         - File</span></span><br><span data-ttu-id="a1120-359">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-359">
         -HtmlCoercion</span></span><br><span data-ttu-id="a1120-360">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-360">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-361">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-361">
         -MatrixBindings</span></span><br><span data-ttu-id="a1120-362">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-362">
         -MatrixCoercion</span></span><br><span data-ttu-id="a1120-363">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-363">
         -OoxmlCoercion</span></span><br><span data-ttu-id="a1120-364">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-364">
         -PdfFile</span></span><br><span data-ttu-id="a1120-365">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-365">
         - Selection</span></span><br><span data-ttu-id="a1120-366">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-366">
         - Settings</span></span><br><span data-ttu-id="a1120-367">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-367">
         -TableBindings</span></span><br><span data-ttu-id="a1120-368">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-368">
         -TableCoercion</span></span><br><span data-ttu-id="a1120-369">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-369">
         -TextBindings</span></span><br><span data-ttu-id="a1120-370">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-370">
         -TextCoercion</span></span><br><span data-ttu-id="a1120-371">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a1120-371">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-372">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="a1120-372">Office for iOS</span></span></td>
    <td> <span data-ttu-id="a1120-373">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-373">- Taskpane</span></span></td>
    <td> <span data-ttu-id="a1120-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a1120-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a1120-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a1120-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a1120-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a1120-378">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-378">-BindingEvents</span></span><br><span data-ttu-id="a1120-379">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-379">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-380">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a1120-380">customXmlParts</span></span><br><span data-ttu-id="a1120-381">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-381">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-382">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-382">
         - File</span></span><br><span data-ttu-id="a1120-383">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-383">
         -HtmlCoercion</span></span><br><span data-ttu-id="a1120-384">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-384">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-385">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-385">
         -MatrixBindings</span></span><br><span data-ttu-id="a1120-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="a1120-387">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-387">
         -OoxmlCoercion</span></span><br><span data-ttu-id="a1120-388">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-388">
         -PdfFile</span></span><br><span data-ttu-id="a1120-389">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-389">
         - Selection</span></span><br><span data-ttu-id="a1120-390">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-390">
         - Settings</span></span><br><span data-ttu-id="a1120-391">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-391">
         -TableBindings</span></span><br><span data-ttu-id="a1120-392">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-392">
         -TableCoercion</span></span><br><span data-ttu-id="a1120-393">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-393">
         -TextBindings</span></span><br><span data-ttu-id="a1120-394">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-394">
         -TextCoercion</span></span><br><span data-ttu-id="a1120-395">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a1120-395">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-396">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="a1120-396">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="a1120-397">- Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-397">- Taskpane</span></span><br><span data-ttu-id="a1120-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a1120-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a1120-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a1120-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a1120-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a1120-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a1120-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a1120-403">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-403">-BindingEvents</span></span><br><span data-ttu-id="a1120-404">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-404">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-405">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a1120-405">customXmlParts</span></span><br><span data-ttu-id="a1120-406">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-406">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-407">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-407">
         - File</span></span><br><span data-ttu-id="a1120-408">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-408">
         -HtmlCoercion</span></span><br><span data-ttu-id="a1120-409">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-409">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-410">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-410">
         -MatrixBindings</span></span><br><span data-ttu-id="a1120-411">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-411">
         -MatrixCoercion</span></span><br><span data-ttu-id="a1120-412">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-412">
         -OoxmlCoercion</span></span><br><span data-ttu-id="a1120-413">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-413">
         -PdfFile</span></span><br><span data-ttu-id="a1120-414">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-414">
         - Selection</span></span><br><span data-ttu-id="a1120-415">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-415">
         - Settings</span></span><br><span data-ttu-id="a1120-416">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-416">
         -TableBindings</span></span><br><span data-ttu-id="a1120-417">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-417">
         -TableCoercion</span></span><br><span data-ttu-id="a1120-418">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a1120-418">
         -TextBindings</span></span><br><span data-ttu-id="a1120-419">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-419">
         -TextCoercion</span></span><br><span data-ttu-id="a1120-420">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a1120-420">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="a1120-421">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a1120-421">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a1120-422">Платформа</span><span class="sxs-lookup"><span data-stu-id="a1120-422">Platform</span></span></th>
    <th><span data-ttu-id="a1120-423">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="a1120-423">Extension points</span></span></th>
    <th><span data-ttu-id="a1120-424">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="a1120-424">API requirement sets</span></span></th>
    <th><span data-ttu-id="a1120-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="a1120-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-426">Office Online</span><span class="sxs-lookup"><span data-stu-id="a1120-426">Office Online</span></span></td>
    <td> <span data-ttu-id="a1120-427">- Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-427">- Content</span></span><br><span data-ttu-id="a1120-428">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-428">
         - Taskpane</span></span><br><span data-ttu-id="a1120-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-431">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a1120-431">-ActiveView</span></span><br><span data-ttu-id="a1120-432">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-432">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-433">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-433">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-434">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-434">
         - File</span></span><br><span data-ttu-id="a1120-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-435">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-436">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-436">
         -PdfFile</span></span><br><span data-ttu-id="a1120-437">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-437">
         - Selection</span></span><br><span data-ttu-id="a1120-438">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-438">
         - Settings</span></span><br><span data-ttu-id="a1120-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-439">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-440">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-440">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="a1120-441">- Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-441">- Content</span></span><br><span data-ttu-id="a1120-442">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-442">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="a1120-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a1120-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a1120-444">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a1120-444">-ActiveView</span></span><br><span data-ttu-id="a1120-445">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-445">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-446">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-446">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-447">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-447">
         - File</span></span><br><span data-ttu-id="a1120-448">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-448">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-449">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-449">
         -PdfFile</span></span><br><span data-ttu-id="a1120-450">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-450">
         - Selection</span></span><br><span data-ttu-id="a1120-451">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-451">
         - Settings</span></span><br><span data-ttu-id="a1120-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-452">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-453">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="a1120-453">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="a1120-454">- Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-454">- Content</span></span><br><span data-ttu-id="a1120-455">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-455">
         - Taskpane</span></span><br><span data-ttu-id="a1120-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-458">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a1120-458">-ActiveView</span></span><br><span data-ttu-id="a1120-459">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-459">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-460">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-460">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-461">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-461">
         - File</span></span><br><span data-ttu-id="a1120-462">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-462">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-463">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-463">
         -PdfFile</span></span><br><span data-ttu-id="a1120-464">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-464">
         - Selection</span></span><br><span data-ttu-id="a1120-465">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-465">
         - Settings</span></span><br><span data-ttu-id="a1120-466">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-466">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-467">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="a1120-467">Office for iOS</span></span></td>
    <td> <span data-ttu-id="a1120-468">- Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-468">- Content</span></span><br><span data-ttu-id="a1120-469">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-469">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="a1120-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="a1120-471">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a1120-471">-ActiveView</span></span><br><span data-ttu-id="a1120-472">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-472">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-473">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-473">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-474">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-474">
         - File</span></span><br><span data-ttu-id="a1120-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-475">
         -PdfFile</span></span><br><span data-ttu-id="a1120-476">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-476">
         - Selection</span></span><br><span data-ttu-id="a1120-477">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-477">
         - Settings</span></span><br><span data-ttu-id="a1120-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-478">
         -TextCoercion</span></span><br><span data-ttu-id="a1120-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-479">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-480">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="a1120-480">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="a1120-481">- Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-481">- Content</span></span><br><span data-ttu-id="a1120-482">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-482">
         - Taskpane</span></span><br><span data-ttu-id="a1120-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-485">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a1120-485">-ActiveView</span></span><br><span data-ttu-id="a1120-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a1120-486">
         -CompressedFile</span></span><br><span data-ttu-id="a1120-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-487">
         -DocumentEvents</span></span><br><span data-ttu-id="a1120-488">
         - File</span><span class="sxs-lookup"><span data-stu-id="a1120-488">
         - File</span></span><br><span data-ttu-id="a1120-489">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-489">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-490">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a1120-490">
         -PdfFile</span></span><br><span data-ttu-id="a1120-491">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a1120-491">
         - Selection</span></span><br><span data-ttu-id="a1120-492">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-492">
         - Settings</span></span><br><span data-ttu-id="a1120-493">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-493">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="a1120-494">OneNote</span><span class="sxs-lookup"><span data-stu-id="a1120-494">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a1120-495">Платформа</span><span class="sxs-lookup"><span data-stu-id="a1120-495">Platform</span></span></th>
    <th><span data-ttu-id="a1120-496">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="a1120-496">Extension points</span></span></th>
    <th><span data-ttu-id="a1120-497">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="a1120-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="a1120-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="a1120-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="a1120-499">Office Online</span><span class="sxs-lookup"><span data-stu-id="a1120-499">Office Online</span></span></td>
    <td> <span data-ttu-id="a1120-500">- Контент</span><span class="sxs-lookup"><span data-stu-id="a1120-500">- Content</span></span><br><span data-ttu-id="a1120-501">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="a1120-501">
         - Taskpane</span></span><br><span data-ttu-id="a1120-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="a1120-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a1120-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="a1120-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a1120-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a1120-505">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a1120-505">-DocumentEvents</span></span><br><span data-ttu-id="a1120-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-506">
         -HtmlCoercion</span></span><br><span data-ttu-id="a1120-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-507">
         -ImageCoercion</span></span><br><span data-ttu-id="a1120-508">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a1120-508">
         - Settings</span></span><br><span data-ttu-id="a1120-509">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a1120-509">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="a1120-510">См. также</span><span class="sxs-lookup"><span data-stu-id="a1120-510">See also</span></span>

- [<span data-ttu-id="a1120-511">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a1120-511">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="a1120-512">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="a1120-512">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="a1120-513">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="a1120-513">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="a1120-514">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="a1120-514">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
