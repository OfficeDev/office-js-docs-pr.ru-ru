---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: a3e9c508a5bae0e7eb660458835b9242d0602818
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199615"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9e324-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9e324-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9e324-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="9e324-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="9e324-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="9e324-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="9e324-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="9e324-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="9e324-108">Excel</span><span class="sxs-lookup"><span data-stu-id="9e324-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9e324-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="9e324-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9e324-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9e324-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9e324-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9e324-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9e324-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9e324-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="9e324-113">Office Online</span></span></td>
    <td> <span data-ttu-id="9e324-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-114">- TaskPane</span></span><br><span data-ttu-id="9e324-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-115">
        - Content</span></span><br><span data-ttu-id="9e324-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="9e324-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9e324-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9e324-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9e324-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9e324-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9e324-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9e324-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9e324-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9e324-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9e324-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9e324-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9e324-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-126">
        - BindingEvents</span></span><br><span data-ttu-id="9e324-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-127">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-128">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-129">
        - File</span></span><br><span data-ttu-id="9e324-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-130">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-132">
        - Selection</span></span><br><span data-ttu-id="9e324-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-133">
        - Settings</span></span><br><span data-ttu-id="9e324-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-134">
        - TableBindings</span></span><br><span data-ttu-id="9e324-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-135">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-136">
        - TextBindings</span></span><br><span data-ttu-id="9e324-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="9e324-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-139">
        - TaskPane</span></span><br><span data-ttu-id="9e324-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9e324-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9e324-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="9e324-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-142">
        - BindingEvents</span></span><br><span data-ttu-id="9e324-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-143">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-144">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-145">
        - File</span></span><br><span data-ttu-id="9e324-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-146">
        - ImageCoercion</span></span><br><span data-ttu-id="9e324-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-147">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-149">
        - Selection</span></span><br><span data-ttu-id="9e324-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-150">
        - Settings</span></span><br><span data-ttu-id="9e324-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-151">
        - TableBindings</span></span><br><span data-ttu-id="9e324-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-152">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-153">
        - TextBindings</span></span><br><span data-ttu-id="9e324-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="9e324-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-156">- TaskPane</span></span><br><span data-ttu-id="9e324-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-157">
        - Content</span></span></td>
    <td><span data-ttu-id="9e324-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9e324-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9e324-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="9e324-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-160">- BindingEvents</span></span><br><span data-ttu-id="9e324-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-161">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-162">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-163">
        - File</span></span><br><span data-ttu-id="9e324-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-164">
        - ImageCoercion</span></span><br><span data-ttu-id="9e324-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-165">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-167">
        - Selection</span></span><br><span data-ttu-id="9e324-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-168">
        - Settings</span></span><br><span data-ttu-id="9e324-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-169">
        - TableBindings</span></span><br><span data-ttu-id="9e324-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-170">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-171">
        - TextBindings</span></span><br><span data-ttu-id="9e324-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="9e324-174">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-174">- TaskPane</span></span><br><span data-ttu-id="9e324-175">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-175">
        - Content</span></span><br><span data-ttu-id="9e324-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9e324-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9e324-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9e324-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9e324-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9e324-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9e324-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9e324-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9e324-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9e324-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9e324-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9e324-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-186">- BindingEvents</span></span><br><span data-ttu-id="9e324-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-187">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-188">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-189">
        - File</span></span><br><span data-ttu-id="9e324-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-190">
        - ImageCoercion</span></span><br><span data-ttu-id="9e324-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-191">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-193">
        - Selection</span></span><br><span data-ttu-id="9e324-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-194">
        - Settings</span></span><br><span data-ttu-id="9e324-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-195">
        - TableBindings</span></span><br><span data-ttu-id="9e324-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-196">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-197">
        - TextBindings</span></span><br><span data-ttu-id="9e324-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-199">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9e324-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="9e324-200">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-200">- TaskPane</span></span><br><span data-ttu-id="9e324-201">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-201">
        - Content</span></span></td>
    <td><span data-ttu-id="9e324-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9e324-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9e324-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9e324-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9e324-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9e324-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9e324-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9e324-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9e324-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9e324-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9e324-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-211">- BindingEvents</span></span><br><span data-ttu-id="9e324-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-212">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-213">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-214">
        - File</span></span><br><span data-ttu-id="9e324-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-215">
        - ImageCoercion</span></span><br><span data-ttu-id="9e324-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-216">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-218">
        - Selection</span></span><br><span data-ttu-id="9e324-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-219">
        - Settings</span></span><br><span data-ttu-id="9e324-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-220">
        - TableBindings</span></span><br><span data-ttu-id="9e324-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-221">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-222">
        - TextBindings</span></span><br><span data-ttu-id="9e324-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-224">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="9e324-225">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-225">- TaskPane</span></span><br><span data-ttu-id="9e324-226">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-226">
        - Content</span></span></td>
    <td><span data-ttu-id="9e324-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9e324-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9e324-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="9e324-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-229">- BindingEvents</span></span><br><span data-ttu-id="9e324-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-230">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-231">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-232">
        - File</span></span><br><span data-ttu-id="9e324-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-233">
        - ImageCoercion</span></span><br><span data-ttu-id="9e324-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-234">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-236">
        - PdfFile</span></span><br><span data-ttu-id="9e324-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-237">
        - Selection</span></span><br><span data-ttu-id="9e324-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-238">
        - Settings</span></span><br><span data-ttu-id="9e324-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-239">
        - TableBindings</span></span><br><span data-ttu-id="9e324-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-240">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-241">
        - TextBindings</span></span><br><span data-ttu-id="9e324-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-243">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="9e324-244">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-244">- TaskPane</span></span><br><span data-ttu-id="9e324-245">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-245">
        - Content</span></span><br><span data-ttu-id="9e324-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9e324-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9e324-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9e324-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9e324-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9e324-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9e324-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9e324-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="9e324-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="9e324-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="9e324-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9e324-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-256">- BindingEvents</span></span><br><span data-ttu-id="9e324-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-257">
        - CompressedFile</span></span><br><span data-ttu-id="9e324-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-258">
        - DocumentEvents</span></span><br><span data-ttu-id="9e324-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="9e324-259">
        - File</span></span><br><span data-ttu-id="9e324-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-260">
        - ImageCoercion</span></span><br><span data-ttu-id="9e324-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-261">
        - MatrixBindings</span></span><br><span data-ttu-id="9e324-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="9e324-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-263">
        - PdfFile</span></span><br><span data-ttu-id="9e324-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-264">
        - Selection</span></span><br><span data-ttu-id="9e324-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-265">
        - Settings</span></span><br><span data-ttu-id="9e324-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-266">
        - TableBindings</span></span><br><span data-ttu-id="9e324-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-267">
        - TableCoercion</span></span><br><span data-ttu-id="9e324-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-268">
        - TextBindings</span></span><br><span data-ttu-id="9e324-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="9e324-270">Outlook</span><span class="sxs-lookup"><span data-stu-id="9e324-270">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9e324-271">Платформа</span><span class="sxs-lookup"><span data-stu-id="9e324-271">Platform</span></span></th>
    <th><span data-ttu-id="9e324-272">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9e324-272">Extension points</span></span></th>
    <th><span data-ttu-id="9e324-273">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9e324-273">API requirement sets</span></span></th>
    <th><span data-ttu-id="9e324-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9e324-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-275">Office Online</span><span class="sxs-lookup"><span data-stu-id="9e324-275">Office Online</span></span></td>
    <td> <span data-ttu-id="9e324-276">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-276">- Mail Read</span></span><br><span data-ttu-id="9e324-277">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9e324-277">
      - Mail Compose</span></span><br><span data-ttu-id="9e324-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9e324-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9e324-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9e324-286">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-286">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-287">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-288">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-288">- Mail Read</span></span><br><span data-ttu-id="9e324-289">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9e324-289">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="9e324-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="9e324-294">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-294">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-295">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-295">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-296">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-296">- Mail Read</span></span><br><span data-ttu-id="9e324-297">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9e324-297">
      - Mail Compose</span></span><br><span data-ttu-id="9e324-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9e324-299">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="9e324-299">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9e324-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9e324-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9e324-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9e324-307">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-307">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-308">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-308">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-309">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-309">- Mail Read</span></span><br><span data-ttu-id="9e324-310">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9e324-310">
      - Mail Compose</span></span><br><span data-ttu-id="9e324-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9e324-312">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="9e324-312">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9e324-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9e324-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9e324-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9e324-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9e324-320">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-320">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-321">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="9e324-321">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9e324-322">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-322">- Mail Read</span></span><br><span data-ttu-id="9e324-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9e324-329">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-329">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-330">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-330">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9e324-331">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-331">- Mail Read</span></span><br><span data-ttu-id="9e324-332">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9e324-332">
      - Mail Compose</span></span><br><span data-ttu-id="9e324-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9e324-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9e324-340">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-341">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-341">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="9e324-342">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-342">- Mail Read</span></span><br><span data-ttu-id="9e324-343">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="9e324-343">
      - Mail Compose</span></span><br><span data-ttu-id="9e324-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9e324-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9e324-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9e324-351">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-352">Office для Android</span><span class="sxs-lookup"><span data-stu-id="9e324-352">Office for Android</span></span></td>
    <td> <span data-ttu-id="9e324-353">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="9e324-353">- Mail Read</span></span><br><span data-ttu-id="9e324-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9e324-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9e324-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9e324-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9e324-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9e324-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9e324-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9e324-360">Недоступно</span><span class="sxs-lookup"><span data-stu-id="9e324-360">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="9e324-361">Word</span><span class="sxs-lookup"><span data-stu-id="9e324-361">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9e324-362">Платформа</span><span class="sxs-lookup"><span data-stu-id="9e324-362">Platform</span></span></th>
    <th><span data-ttu-id="9e324-363">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9e324-363">Extension points</span></span></th>
    <th><span data-ttu-id="9e324-364">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9e324-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="9e324-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9e324-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-366">Office Online</span><span class="sxs-lookup"><span data-stu-id="9e324-366">Office Online</span></span></td>
    <td> <span data-ttu-id="9e324-367">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-367">- TaskPane</span></span><br><span data-ttu-id="9e324-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9e324-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9e324-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9e324-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-373">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-373">- BindingEvents</span></span><br><span data-ttu-id="9e324-374">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-374">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-375">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-376">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-376">
         - File</span></span><br><span data-ttu-id="9e324-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-377">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-378">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-379">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-379">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-380">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-380">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-381">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-381">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-382">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-382">
         - PdfFile</span></span><br><span data-ttu-id="9e324-383">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-383">
         - Selection</span></span><br><span data-ttu-id="9e324-384">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-384">
         - Settings</span></span><br><span data-ttu-id="9e324-385">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-385">
         - TableBindings</span></span><br><span data-ttu-id="9e324-386">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-386">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-387">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-387">
         - TextBindings</span></span><br><span data-ttu-id="9e324-388">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-388">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-389">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-389">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-390">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-390">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-391">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-391">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9e324-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9e324-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-393">- BindingEvents</span></span><br><span data-ttu-id="9e324-394">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-394">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-395">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-395">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-396">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-396">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-397">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-397">
         - File</span></span><br><span data-ttu-id="9e324-398">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-398">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-399">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-399">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-400">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-400">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-401">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-401">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-402">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-402">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-403">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-403">
         - PdfFile</span></span><br><span data-ttu-id="9e324-404">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-404">
         - Selection</span></span><br><span data-ttu-id="9e324-405">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-405">
         - Settings</span></span><br><span data-ttu-id="9e324-406">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-406">
         - TableBindings</span></span><br><span data-ttu-id="9e324-407">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-407">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-408">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-408">
         - TextBindings</span></span><br><span data-ttu-id="9e324-409">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-409">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-410">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-410">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-411">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-411">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-412">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-412">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9e324-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9e324-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="9e324-415">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-415">- BindingEvents</span></span><br><span data-ttu-id="9e324-416">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-416">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-417">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-417">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-418">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-418">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-419">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-419">
         - File</span></span><br><span data-ttu-id="9e324-420">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-420">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-421">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-421">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-422">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-422">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-423">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-423">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-424">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-424">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-425">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-425">
         - PdfFile</span></span><br><span data-ttu-id="9e324-426">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-426">
         - Selection</span></span><br><span data-ttu-id="9e324-427">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-427">
         - Settings</span></span><br><span data-ttu-id="9e324-428">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-428">
         - TableBindings</span></span><br><span data-ttu-id="9e324-429">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-429">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-430">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-430">
         - TextBindings</span></span><br><span data-ttu-id="9e324-431">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-431">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-432">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-432">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-433">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-433">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-434">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-434">- TaskPane</span></span><br><span data-ttu-id="9e324-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9e324-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9e324-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9e324-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-440">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-440">- BindingEvents</span></span><br><span data-ttu-id="9e324-441">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-441">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-442">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-442">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-443">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-443">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-444">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-444">
         - File</span></span><br><span data-ttu-id="9e324-445">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-445">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-446">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-446">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-447">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-447">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-448">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-448">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-449">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-449">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-450">
         - PdfFile</span></span><br><span data-ttu-id="9e324-451">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-451">
         - Selection</span></span><br><span data-ttu-id="9e324-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-452">
         - Settings</span></span><br><span data-ttu-id="9e324-453">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-453">
         - TableBindings</span></span><br><span data-ttu-id="9e324-454">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-454">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-455">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-455">
         - TextBindings</span></span><br><span data-ttu-id="9e324-456">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-456">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-457">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-457">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-458">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9e324-458">Office for iPad</span></span></td>
    <td> <span data-ttu-id="9e324-459">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-459">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9e324-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9e324-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9e324-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9e324-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9e324-464">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-464">- BindingEvents</span></span><br><span data-ttu-id="9e324-465">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-465">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-466">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-466">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-467">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-467">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-468">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-468">
         - File</span></span><br><span data-ttu-id="9e324-469">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-469">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-470">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-470">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-471">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-471">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-472">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-472">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-473">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-473">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-474">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-474">
         - PdfFile</span></span><br><span data-ttu-id="9e324-475">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-475">
         - Selection</span></span><br><span data-ttu-id="9e324-476">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-476">
         - Settings</span></span><br><span data-ttu-id="9e324-477">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-477">
         - TableBindings</span></span><br><span data-ttu-id="9e324-478">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-478">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-479">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-479">
         - TextBindings</span></span><br><span data-ttu-id="9e324-480">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-480">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-481">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-481">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-482">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-482">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9e324-483">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-483">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9e324-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="9e324-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="9e324-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-486">- BindingEvents</span></span><br><span data-ttu-id="9e324-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-487">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-488">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-489">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-490">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-490">
         - File</span></span><br><span data-ttu-id="9e324-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-491">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-492">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-493">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-494">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-495">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-496">
         - PdfFile</span></span><br><span data-ttu-id="9e324-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-497">
         - Selection</span></span><br><span data-ttu-id="9e324-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-498">
         - Settings</span></span><br><span data-ttu-id="9e324-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-499">
         - TableBindings</span></span><br><span data-ttu-id="9e324-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-500">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-501">
         - TextBindings</span></span><br><span data-ttu-id="9e324-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-502">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-503">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-504">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-504">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="9e324-505">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-505">- TaskPane</span></span><br><span data-ttu-id="9e324-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9e324-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9e324-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9e324-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9e324-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9e324-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9e324-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9e324-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-511">- BindingEvents</span></span><br><span data-ttu-id="9e324-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-512">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="9e324-513">
         - CustomXmlParts</span></span><br><span data-ttu-id="9e324-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-514">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-515">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="9e324-515">
         - File</span></span><br><span data-ttu-id="9e324-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-516">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-517">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-518">
         - MatrixBindings</span></span><br><span data-ttu-id="9e324-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-519">
         - MatrixCoercion</span></span><br><span data-ttu-id="9e324-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-520">
         - OoxmlCoercion</span></span><br><span data-ttu-id="9e324-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-521">
         - PdfFile</span></span><br><span data-ttu-id="9e324-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-522">
         - Selection</span></span><br><span data-ttu-id="9e324-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="9e324-523">
         - Settings</span></span><br><span data-ttu-id="9e324-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-524">
         - TableBindings</span></span><br><span data-ttu-id="9e324-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-525">
         - TableCoercion</span></span><br><span data-ttu-id="9e324-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9e324-526">
         - TextBindings</span></span><br><span data-ttu-id="9e324-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-527">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9e324-528">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9e324-529">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9e324-529">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9e324-530">Платформа</span><span class="sxs-lookup"><span data-stu-id="9e324-530">Platform</span></span></th>
    <th><span data-ttu-id="9e324-531">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9e324-531">Extension points</span></span></th>
    <th><span data-ttu-id="9e324-532">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9e324-532">API requirement sets</span></span></th>
    <th><span data-ttu-id="9e324-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9e324-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-534">Office Online</span><span class="sxs-lookup"><span data-stu-id="9e324-534">Office Online</span></span></td>
    <td> <span data-ttu-id="9e324-535">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-535">- Content</span></span><br><span data-ttu-id="9e324-536">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-536">
         - TaskPane</span></span><br><span data-ttu-id="9e324-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-539">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-539">- ActiveView</span></span><br><span data-ttu-id="9e324-540">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-540">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-541">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-542">
         - File</span></span><br><span data-ttu-id="9e324-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-543">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-544">
         - PdfFile</span></span><br><span data-ttu-id="9e324-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-545">
         - Selection</span></span><br><span data-ttu-id="9e324-546">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-546">
         - Settings</span></span><br><span data-ttu-id="9e324-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-547">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-548">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-548">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-549">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-549">- Content</span></span><br><span data-ttu-id="9e324-550">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-550">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="9e324-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9e324-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9e324-552">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-552">- ActiveView</span></span><br><span data-ttu-id="9e324-553">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-553">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-554">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-554">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-555">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-555">
         - File</span></span><br><span data-ttu-id="9e324-556">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-556">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-557">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-557">
         - PdfFile</span></span><br><span data-ttu-id="9e324-558">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-558">
         - Selection</span></span><br><span data-ttu-id="9e324-559">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-559">
         - Settings</span></span><br><span data-ttu-id="9e324-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-560">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-561">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-561">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-562">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-562">- Content</span></span><br><span data-ttu-id="9e324-563">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-563">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9e324-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9e324-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-565">- ActiveView</span></span><br><span data-ttu-id="9e324-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-566">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-567">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-568">
         - File</span></span><br><span data-ttu-id="9e324-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-569">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-570">
         - PdfFile</span></span><br><span data-ttu-id="9e324-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-571">
         - Selection</span></span><br><span data-ttu-id="9e324-572">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-572">
         - Settings</span></span><br><span data-ttu-id="9e324-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-573">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-574">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-574">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-575">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-575">- Content</span></span><br><span data-ttu-id="9e324-576">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-576">
         - TaskPane</span></span><br><span data-ttu-id="9e324-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-579">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-579">- ActiveView</span></span><br><span data-ttu-id="9e324-580">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-580">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-581">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-581">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-582">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-582">
         - File</span></span><br><span data-ttu-id="9e324-583">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-583">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-584">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-584">
         - PdfFile</span></span><br><span data-ttu-id="9e324-585">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-585">
         - Selection</span></span><br><span data-ttu-id="9e324-586">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-586">
         - Settings</span></span><br><span data-ttu-id="9e324-587">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-587">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-588">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="9e324-588">Office for iPad</span></span></td>
    <td> <span data-ttu-id="9e324-589">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-589">- Content</span></span><br><span data-ttu-id="9e324-590">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-590">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="9e324-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-592">- ActiveView</span></span><br><span data-ttu-id="9e324-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-593">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-594">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-595">
         - File</span></span><br><span data-ttu-id="9e324-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-596">
         - PdfFile</span></span><br><span data-ttu-id="9e324-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-597">
         - Selection</span></span><br><span data-ttu-id="9e324-598">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-598">
         - Settings</span></span><br><span data-ttu-id="9e324-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-599">
         - TextCoercion</span></span><br><span data-ttu-id="9e324-600">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-600">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-601">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-601">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9e324-602">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-602">- Content</span></span><br><span data-ttu-id="9e324-603">
         - Область задач/td></span><span class="sxs-lookup"><span data-stu-id="9e324-603">
         - TaskPane/td></span></span> <td> <span data-ttu-id="9e324-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="9e324-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="9e324-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-605">- ActiveView</span></span><br><span data-ttu-id="9e324-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-606">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-607">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-608">
         - File</span></span><br><span data-ttu-id="9e324-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-609">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-610">
         - PdfFile</span></span><br><span data-ttu-id="9e324-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-611">
         - Selection</span></span><br><span data-ttu-id="9e324-612">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-612">
         - Settings</span></span><br><span data-ttu-id="9e324-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-613">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-614">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="9e324-614">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="9e324-615">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-615">- Content</span></span><br><span data-ttu-id="9e324-616">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-616">
         - TaskPane</span></span><br><span data-ttu-id="9e324-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9e324-619">- ActiveView</span></span><br><span data-ttu-id="9e324-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9e324-620">
         - CompressedFile</span></span><br><span data-ttu-id="9e324-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-621">
         - DocumentEvents</span></span><br><span data-ttu-id="9e324-622">
         - File</span><span class="sxs-lookup"><span data-stu-id="9e324-622">
         - File</span></span><br><span data-ttu-id="9e324-623">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-623">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9e324-624">
         - PdfFile</span></span><br><span data-ttu-id="9e324-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-625">
         - Selection</span></span><br><span data-ttu-id="9e324-626">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-626">
         - Settings</span></span><br><span data-ttu-id="9e324-627">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-627">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="9e324-628">OneNote</span><span class="sxs-lookup"><span data-stu-id="9e324-628">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9e324-629">Платформа</span><span class="sxs-lookup"><span data-stu-id="9e324-629">Platform</span></span></th>
    <th><span data-ttu-id="9e324-630">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9e324-630">Extension points</span></span></th>
    <th><span data-ttu-id="9e324-631">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9e324-631">API requirement sets</span></span></th>
    <th><span data-ttu-id="9e324-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9e324-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-633">Office Online</span><span class="sxs-lookup"><span data-stu-id="9e324-633">Office Online</span></span></td>
    <td> <span data-ttu-id="9e324-634">- Контент</span><span class="sxs-lookup"><span data-stu-id="9e324-634">- Content</span></span><br><span data-ttu-id="9e324-635">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-635">
         - TaskPane</span></span><br><span data-ttu-id="9e324-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="9e324-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9e324-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9e324-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-639">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9e324-639">- DocumentEvents</span></span><br><span data-ttu-id="9e324-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="9e324-641">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-641">
         - ImageCoercion</span></span><br><span data-ttu-id="9e324-642">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="9e324-642">
         - Settings</span></span><br><span data-ttu-id="9e324-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-643">
         - TextCoercion</span></span></td>
  </tr>
</table><span data-ttu-id="9e324-644">
\*&ast; - Добавлены обновления после выпуска.*

</span><span class="sxs-lookup"><span data-stu-id="9e324-644">
\*&ast; - Added with post-release updates.*

</span></span><br/>

## <a name="project"></a><span data-ttu-id="9e324-645">Project</span><span class="sxs-lookup"><span data-stu-id="9e324-645">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9e324-646">Платформа</span><span class="sxs-lookup"><span data-stu-id="9e324-646">Platform</span></span></th>
    <th><span data-ttu-id="9e324-647">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="9e324-647">Extension points</span></span></th>
    <th><span data-ttu-id="9e324-648">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="9e324-648">API requirement sets</span></span></th>
    <th><span data-ttu-id="9e324-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="9e324-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-650">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-650">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-651">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-653">- Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-653">- Selection</span></span><br><span data-ttu-id="9e324-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-654">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-655">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-655">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-656">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-656">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-658">- Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-658">- Selection</span></span><br><span data-ttu-id="9e324-659">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-659">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9e324-660">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="9e324-660">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="9e324-661">- Область задач</span><span class="sxs-lookup"><span data-stu-id="9e324-661">- TaskPane</span></span></td>
    <td> <span data-ttu-id="9e324-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9e324-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9e324-663">- Selection</span><span class="sxs-lookup"><span data-stu-id="9e324-663">- Selection</span></span><br><span data-ttu-id="9e324-664">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9e324-664">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9e324-665">См. также</span><span class="sxs-lookup"><span data-stu-id="9e324-665">See also</span></span>

- [<span data-ttu-id="9e324-666">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9e324-666">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9e324-667">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="9e324-667">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="9e324-668">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="9e324-668">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="9e324-669">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="9e324-669">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
