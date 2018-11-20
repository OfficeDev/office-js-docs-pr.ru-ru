---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 11/07/2018
ms.openlocfilehash: f8d7d9d393531301829b31dd171a5332a0da536b
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533800"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="511f6-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="511f6-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="511f6-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="511f6-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="511f6-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="511f6-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="511f6-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="511f6-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="511f6-108">Excel</span><span class="sxs-lookup"><span data-stu-id="511f6-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="511f6-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="511f6-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="511f6-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="511f6-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="511f6-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="511f6-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="511f6-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="511f6-113">Office Online</span></span></td>
    <td> <span data-ttu-id="511f6-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-114">- Taskpane</span></span><br><span data-ttu-id="511f6-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-115">
        - Content</span></span><br><span data-ttu-id="511f6-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="511f6-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="511f6-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="511f6-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="511f6-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="511f6-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="511f6-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="511f6-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="511f6-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-123">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="511f6-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="511f6-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="511f6-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-126">
        -BindingEvents</span></span><br><span data-ttu-id="511f6-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-127">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-128">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-129">
        - File</span></span><br><span data-ttu-id="511f6-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-130">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-131">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-132">
        - Selection</span></span><br><span data-ttu-id="511f6-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-133">
        - Settings</span></span><br><span data-ttu-id="511f6-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-134">
        -TableBindings</span></span><br><span data-ttu-id="511f6-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-135">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-136">
        -TextBindings</span></span><br><span data-ttu-id="511f6-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-137">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="511f6-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-139">
        - Taskpane</span></span><br><span data-ttu-id="511f6-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="511f6-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-142">
        -BindingEvents</span></span><br><span data-ttu-id="511f6-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-143">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-144">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-145">
        - File</span></span><br><span data-ttu-id="511f6-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-146">
        -ImageCoercion</span></span><br><span data-ttu-id="511f6-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-147">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-148">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-149">
        - Selection</span></span><br><span data-ttu-id="511f6-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-150">
        - Settings</span></span><br><span data-ttu-id="511f6-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-151">
        -TableBindings</span></span><br><span data-ttu-id="511f6-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-152">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-153">
        -TextBindings</span></span><br><span data-ttu-id="511f6-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-154">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="511f6-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-156">- Taskpane</span></span><br><span data-ttu-id="511f6-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-157">
        - Content</span></span><br><span data-ttu-id="511f6-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="511f6-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="511f6-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="511f6-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="511f6-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="511f6-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="511f6-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="511f6-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-165">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="511f6-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="511f6-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="511f6-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-168">-BindingEvents</span></span><br><span data-ttu-id="511f6-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-169">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-170">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-171">
        - File</span></span><br><span data-ttu-id="511f6-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-172">
        -ImageCoercion</span></span><br><span data-ttu-id="511f6-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-173">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-175">
        - Selection</span></span><br><span data-ttu-id="511f6-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-176">
        - Settings</span></span><br><span data-ttu-id="511f6-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-177">
        -TableBindings</span></span><br><span data-ttu-id="511f6-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-178">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-179">
        -TextBindings</span></span><br><span data-ttu-id="511f6-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-181">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-181">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="511f6-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-182">- Taskpane</span></span><br><span data-ttu-id="511f6-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-183">
        - Content</span></span><br><span data-ttu-id="511f6-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="511f6-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="511f6-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="511f6-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="511f6-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="511f6-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="511f6-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="511f6-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="511f6-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="511f6-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="511f6-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-194">-BindingEvents</span></span><br><span data-ttu-id="511f6-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-195">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-196">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-197">
        - File</span></span><br><span data-ttu-id="511f6-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-198">
        -ImageCoercion</span></span><br><span data-ttu-id="511f6-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-199">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-200">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-201">
        - Selection</span></span><br><span data-ttu-id="511f6-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-202">
        - Settings</span></span><br><span data-ttu-id="511f6-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-203">
        -TableBindings</span></span><br><span data-ttu-id="511f6-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-204">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-205">
        -TextBindings</span></span><br><span data-ttu-id="511f6-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-206">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-207">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="511f6-207">Office for iOS</span></span></td>
    <td><span data-ttu-id="511f6-208">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-208">- Taskpane</span></span><br><span data-ttu-id="511f6-209">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-209">
        - Content</span></span></td>
    <td><span data-ttu-id="511f6-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="511f6-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="511f6-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="511f6-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="511f6-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="511f6-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="511f6-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-216">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="511f6-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="511f6-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="511f6-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-219">-BindingEvents</span></span><br><span data-ttu-id="511f6-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-220">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-221">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-222">
        - File</span></span><br><span data-ttu-id="511f6-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-223">
        -ImageCoercion</span></span><br><span data-ttu-id="511f6-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-224">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-225">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-226">
        - Selection</span></span><br><span data-ttu-id="511f6-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-227">
        - Settings</span></span><br><span data-ttu-id="511f6-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-228">
        -TableBindings</span></span><br><span data-ttu-id="511f6-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-229">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-230">
        -TextBindings</span></span><br><span data-ttu-id="511f6-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-231">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-232">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="511f6-233">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-233">- Taskpane</span></span><br><span data-ttu-id="511f6-234">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-234">
        - Content</span></span><br><span data-ttu-id="511f6-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="511f6-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="511f6-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="511f6-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="511f6-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="511f6-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="511f6-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="511f6-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-242">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="511f6-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="511f6-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="511f6-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-245">-BindingEvents</span></span><br><span data-ttu-id="511f6-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-246">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-247">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-248">
        - File</span></span><br><span data-ttu-id="511f6-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-249">
        -ImageCoercion</span></span><br><span data-ttu-id="511f6-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-250">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-251">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-252">
        -PdfFile</span></span><br><span data-ttu-id="511f6-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-253">
        - Selection</span></span><br><span data-ttu-id="511f6-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-254">
        - Settings</span></span><br><span data-ttu-id="511f6-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-255">
        -TableBindings</span></span><br><span data-ttu-id="511f6-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-256">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-257">
        -TextBindings</span></span><br><span data-ttu-id="511f6-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-258">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-259">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-259">Office for Mac</span></span></td>
    <td><span data-ttu-id="511f6-260">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-260">- Taskpane</span></span><br><span data-ttu-id="511f6-261">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-261">
        - Content</span></span><br><span data-ttu-id="511f6-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="511f6-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="511f6-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="511f6-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="511f6-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="511f6-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="511f6-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="511f6-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-269">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="511f6-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="511f6-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="511f6-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="511f6-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-272">-BindingEvents</span></span><br><span data-ttu-id="511f6-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-273">
        -CompressedFile</span></span><br><span data-ttu-id="511f6-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-274">
        -DocumentEvents</span></span><br><span data-ttu-id="511f6-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="511f6-275">
        - File</span></span><br><span data-ttu-id="511f6-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-276">
        -ImageCoercion</span></span><br><span data-ttu-id="511f6-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-277">
        -MatrixBindings</span></span><br><span data-ttu-id="511f6-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-278">
        -MatrixCoercion</span></span><br><span data-ttu-id="511f6-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-279">
        -PdfFile</span></span><br><span data-ttu-id="511f6-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-280">
        - Selection</span></span><br><span data-ttu-id="511f6-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-281">
        - Settings</span></span><br><span data-ttu-id="511f6-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-282">
        -TableBindings</span></span><br><span data-ttu-id="511f6-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-283">
        -TableCoercion</span></span><br><span data-ttu-id="511f6-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-284">
        -TextBindings</span></span><br><span data-ttu-id="511f6-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-285">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="511f6-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="511f6-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="511f6-287">Платформа</span><span class="sxs-lookup"><span data-stu-id="511f6-287">Platform</span></span></th>
    <th><span data-ttu-id="511f6-288">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="511f6-288">Extension points</span></span></th>
    <th><span data-ttu-id="511f6-289">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="511f6-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="511f6-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="511f6-291">Office Online</span></span></td>
    <td> <span data-ttu-id="511f6-292">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-292">- Mail Read</span></span><br><span data-ttu-id="511f6-293">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="511f6-293">
      - Mail Compose</span></span><br><span data-ttu-id="511f6-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span><br><span data-ttu-id="511f6-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="511f6-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="511f6-302">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-303">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-304">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-304">- Mail Read</span></span><br><span data-ttu-id="511f6-305">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="511f6-305">
      - Mail Compose</span></span><br><span data-ttu-id="511f6-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span></td>
    <td><span data-ttu-id="511f6-311">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-312">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-313">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-313">- Mail Read</span></span><br><span data-ttu-id="511f6-314">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="511f6-314">
      - Mail Compose</span></span><br><span data-ttu-id="511f6-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="511f6-316">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="511f6-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="511f6-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span><br><span data-ttu-id="511f6-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="511f6-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="511f6-324">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-325">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-325">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="511f6-326">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-326">- Mail Read</span></span><br><span data-ttu-id="511f6-327">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="511f6-327">
      - Mail Compose</span></span><br><span data-ttu-id="511f6-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="511f6-329">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="511f6-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="511f6-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span><br><span data-ttu-id="511f6-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="511f6-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="511f6-337">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-338">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="511f6-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="511f6-339">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-339">- Mail Read</span></span><br><span data-ttu-id="511f6-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span></td>
    <td><span data-ttu-id="511f6-346">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-347">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="511f6-348">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-348">- Mail Read</span></span><br><span data-ttu-id="511f6-349">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="511f6-349">
      - Mail Compose</span></span><br><span data-ttu-id="511f6-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span><br><span data-ttu-id="511f6-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="511f6-357">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-358">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-358">Office for Mac</span></span></td>
    <td> <span data-ttu-id="511f6-359">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-359">- Mail Read</span></span><br><span data-ttu-id="511f6-360">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="511f6-360">
      - Mail Compose</span></span><br><span data-ttu-id="511f6-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span><br><span data-ttu-id="511f6-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="511f6-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="511f6-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="511f6-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="511f6-369">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-370">Office для Android</span><span class="sxs-lookup"><span data-stu-id="511f6-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="511f6-371">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="511f6-371">- Mail Read</span></span><br><span data-ttu-id="511f6-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"></a> Mailbox 1.1</span></span><br><span data-ttu-id="511f6-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"></a> Mailbox 1.2</span></span><br><span data-ttu-id="511f6-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"></a> Mailbox 1.3</span></span><br><span data-ttu-id="511f6-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="511f6-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"></a> Mailbox 1.4</span></span><br><span data-ttu-id="511f6-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="511f6-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"></a> Mailbox 1.5</span></span></td>
    <td><span data-ttu-id="511f6-378">Недоступно</span><span class="sxs-lookup"><span data-stu-id="511f6-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="511f6-379">Word</span><span class="sxs-lookup"><span data-stu-id="511f6-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="511f6-380">Платформа</span><span class="sxs-lookup"><span data-stu-id="511f6-380">Platform</span></span></th>
    <th><span data-ttu-id="511f6-381">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="511f6-381">Extension points</span></span></th>
    <th><span data-ttu-id="511f6-382">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="511f6-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="511f6-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="511f6-384">Office Online</span></span></td>
    <td> <span data-ttu-id="511f6-385">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-385">- Taskpane</span></span><br><span data-ttu-id="511f6-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="511f6-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="511f6-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="511f6-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-391">-BindingEvents</span></span><br><span data-ttu-id="511f6-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-392">
         -</span></span><br><span data-ttu-id="511f6-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-393">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-394">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-394">
         - File</span></span><br><span data-ttu-id="511f6-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-395">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-396">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-397">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-398">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-399">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-400">
         -PdfFile</span></span><br><span data-ttu-id="511f6-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-401">
         - Selection</span></span><br><span data-ttu-id="511f6-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-402">
         - Settings</span></span><br><span data-ttu-id="511f6-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-403">
         -TableBindings</span></span><br><span data-ttu-id="511f6-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-404">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-405">
         -TextBindings</span></span><br><span data-ttu-id="511f6-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-406">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-407">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-408">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-409">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-409">- Taskpane</span></span></td>
    <td> <span data-ttu-id="511f6-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-411">-BindingEvents</span></span><br><span data-ttu-id="511f6-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-412">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-413">
         -</span></span><br><span data-ttu-id="511f6-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-414">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-415">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-415">
         - File</span></span><br><span data-ttu-id="511f6-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-416">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-417">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-418">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-419">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-420">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-421">
         -PdfFile</span></span><br><span data-ttu-id="511f6-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-422">
         - Selection</span></span><br><span data-ttu-id="511f6-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-423">
         - Settings</span></span><br><span data-ttu-id="511f6-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-424">
         -TableBindings</span></span><br><span data-ttu-id="511f6-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-425">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-426">
         -TextBindings</span></span><br><span data-ttu-id="511f6-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-427">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-428">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-429">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-430">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-430">- Taskpane</span></span><br><span data-ttu-id="511f6-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="511f6-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="511f6-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="511f6-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-436">-BindingEvents</span></span><br><span data-ttu-id="511f6-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-437">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-438">
         -</span></span><br><span data-ttu-id="511f6-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-439">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-440">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-440">
         - File</span></span><br><span data-ttu-id="511f6-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-441">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-442">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-443">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-444">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-445">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-446">
         -PdfFile</span></span><br><span data-ttu-id="511f6-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-447">
         - Selection</span></span><br><span data-ttu-id="511f6-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-448">
         - Settings</span></span><br><span data-ttu-id="511f6-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-449">
         -TableBindings</span></span><br><span data-ttu-id="511f6-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-450">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-451">
         -TextBindings</span></span><br><span data-ttu-id="511f6-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-452">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-453">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-454">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-454">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="511f6-455">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-455">- Taskpane</span></span><br><span data-ttu-id="511f6-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="511f6-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="511f6-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="511f6-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-461">-BindingEvents</span></span><br><span data-ttu-id="511f6-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-462">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-463">
         -</span></span><br><span data-ttu-id="511f6-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-464">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-465">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-465">
         - File</span></span><br><span data-ttu-id="511f6-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-466">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-467">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-468">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-469">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-470">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-471">
         -PdfFile</span></span><br><span data-ttu-id="511f6-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-472">
         - Selection</span></span><br><span data-ttu-id="511f6-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-473">
         - Settings</span></span><br><span data-ttu-id="511f6-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-474">
         -TableBindings</span></span><br><span data-ttu-id="511f6-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-475">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-476">
         -TextBindings</span></span><br><span data-ttu-id="511f6-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-477">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-478">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-479">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="511f6-479">Office for iOS</span></span></td>
    <td> <span data-ttu-id="511f6-480">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-480">- Taskpane</span></span></td>
    <td> <span data-ttu-id="511f6-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="511f6-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="511f6-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="511f6-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="511f6-484">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="511f6-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-485">-BindingEvents</span></span><br><span data-ttu-id="511f6-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-486">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-487">
         -</span></span><br><span data-ttu-id="511f6-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-488">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-489">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-489">
         - File</span></span><br><span data-ttu-id="511f6-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-490">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-491">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-492">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-493">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-494">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-495">
         -PdfFile</span></span><br><span data-ttu-id="511f6-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-496">
         - Selection</span></span><br><span data-ttu-id="511f6-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-497">
         - Settings</span></span><br><span data-ttu-id="511f6-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-498">
         -TableBindings</span></span><br><span data-ttu-id="511f6-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-499">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-500">
         -TextBindings</span></span><br><span data-ttu-id="511f6-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-501">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-502">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-503">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="511f6-504">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-504">- Taskpane</span></span><br><span data-ttu-id="511f6-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="511f6-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="511f6-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="511f6-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="511f6-509">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="511f6-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-510">-BindingEvents</span></span><br><span data-ttu-id="511f6-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-511">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-512">
         -</span></span><br><span data-ttu-id="511f6-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-513">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-514">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-514">
         - File</span></span><br><span data-ttu-id="511f6-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-515">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-516">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-517">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-518">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-519">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-520">
         -PdfFile</span></span><br><span data-ttu-id="511f6-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-521">
         - Selection</span></span><br><span data-ttu-id="511f6-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-522">
         - Settings</span></span><br><span data-ttu-id="511f6-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-523">
         -TableBindings</span></span><br><span data-ttu-id="511f6-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-524">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-525">
         -TextBindings</span></span><br><span data-ttu-id="511f6-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-526">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-527">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-528">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-528">Office for Mac</span></span></td>
    <td> <span data-ttu-id="511f6-529">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-529">- Taskpane</span></span><br><span data-ttu-id="511f6-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="511f6-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="511f6-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="511f6-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="511f6-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="511f6-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="511f6-534">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="511f6-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-535">-BindingEvents</span></span><br><span data-ttu-id="511f6-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-536">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="511f6-537">
         -</span></span><br><span data-ttu-id="511f6-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-538">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="511f6-539">
         - File</span></span><br><span data-ttu-id="511f6-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-540">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-541">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-542">
         -MatrixBindings</span></span><br><span data-ttu-id="511f6-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-543">
         -MatrixCoercion</span></span><br><span data-ttu-id="511f6-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-544">
         -OoxmlCoercion</span></span><br><span data-ttu-id="511f6-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-545">
         -PdfFile</span></span><br><span data-ttu-id="511f6-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-546">
         - Selection</span></span><br><span data-ttu-id="511f6-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="511f6-547">
         - Settings</span></span><br><span data-ttu-id="511f6-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-548">
         -TableBindings</span></span><br><span data-ttu-id="511f6-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-549">
         -TableCoercion</span></span><br><span data-ttu-id="511f6-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="511f6-550">
         -TextBindings</span></span><br><span data-ttu-id="511f6-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-551">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="511f6-552">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="511f6-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="511f6-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="511f6-554">Платформа</span><span class="sxs-lookup"><span data-stu-id="511f6-554">Platform</span></span></th>
    <th><span data-ttu-id="511f6-555">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="511f6-555">Extension points</span></span></th>
    <th><span data-ttu-id="511f6-556">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="511f6-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="511f6-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="511f6-558">Office Online</span></span></td>
    <td> <span data-ttu-id="511f6-559">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-559">- Content</span></span><br><span data-ttu-id="511f6-560">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-560">
         - Taskpane</span></span><br><span data-ttu-id="511f6-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-563">-</span></span><br><span data-ttu-id="511f6-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-564">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-565">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-566">
         - File</span></span><br><span data-ttu-id="511f6-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-567">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-568">
         -PdfFile</span></span><br><span data-ttu-id="511f6-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-569">
         - Selection</span></span><br><span data-ttu-id="511f6-570">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-570">
         - Settings</span></span><br><span data-ttu-id="511f6-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-571">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-572">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-573">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-573">- Content</span></span><br><span data-ttu-id="511f6-574">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-574">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="511f6-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="511f6-575">DialogApi 1.1</span></span></td>
    <td> <span data-ttu-id="511f6-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-576">-</span></span><br><span data-ttu-id="511f6-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-577">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-578">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-579">
         - File</span></span><br><span data-ttu-id="511f6-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-580">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-581">
         -PdfFile</span></span><br><span data-ttu-id="511f6-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-582">
         - Selection</span></span><br><span data-ttu-id="511f6-583">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-583">
         - Settings</span></span><br><span data-ttu-id="511f6-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-584">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-586">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-586">- Content</span></span><br><span data-ttu-id="511f6-587">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-587">
         - Taskpane</span></span><br><span data-ttu-id="511f6-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-590">-</span></span><br><span data-ttu-id="511f6-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-591">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-592">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-593">
         - File</span></span><br><span data-ttu-id="511f6-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-594">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-595">
         -PdfFile</span></span><br><span data-ttu-id="511f6-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-596">
         - Selection</span></span><br><span data-ttu-id="511f6-597">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-597">
         - Settings</span></span><br><span data-ttu-id="511f6-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-598">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-599">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-599">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="511f6-600">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-600">- Content</span></span><br><span data-ttu-id="511f6-601">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-601">
         - Taskpane</span></span><br><span data-ttu-id="511f6-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-604">-</span></span><br><span data-ttu-id="511f6-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-605">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-606">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-607">
         - File</span></span><br><span data-ttu-id="511f6-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-608">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-609">
         -PdfFile</span></span><br><span data-ttu-id="511f6-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-610">
         - Selection</span></span><br><span data-ttu-id="511f6-611">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-611">
         - Settings</span></span><br><span data-ttu-id="511f6-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-612">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-613">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="511f6-613">Office for iOS</span></span></td>
    <td> <span data-ttu-id="511f6-614">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-614">- Content</span></span><br><span data-ttu-id="511f6-615">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-615">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="511f6-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="511f6-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-617">-</span></span><br><span data-ttu-id="511f6-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-618">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-619">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-620">
         - File</span></span><br><span data-ttu-id="511f6-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-621">
         -PdfFile</span></span><br><span data-ttu-id="511f6-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-622">
         - Selection</span></span><br><span data-ttu-id="511f6-623">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-623">
         - Settings</span></span><br><span data-ttu-id="511f6-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-624">
         -TextCoercion</span></span><br><span data-ttu-id="511f6-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-625">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-626">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="511f6-627">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-627">- Content</span></span><br><span data-ttu-id="511f6-628">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-628">
         - Taskpane</span></span><br><span data-ttu-id="511f6-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-631">-</span></span><br><span data-ttu-id="511f6-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-632">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-633">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-634">
         - File</span></span><br><span data-ttu-id="511f6-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-635">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-636">
         -PdfFile</span></span><br><span data-ttu-id="511f6-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-637">
         - Selection</span></span><br><span data-ttu-id="511f6-638">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-638">
         - Settings</span></span><br><span data-ttu-id="511f6-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-639">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-640">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="511f6-640">Office for Mac</span></span></td>
    <td> <span data-ttu-id="511f6-641">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-641">- Content</span></span><br><span data-ttu-id="511f6-642">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-642">
         - Taskpane</span></span><br><span data-ttu-id="511f6-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="511f6-645">-</span></span><br><span data-ttu-id="511f6-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="511f6-646">
         -CompressedFile</span></span><br><span data-ttu-id="511f6-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-647">
         -DocumentEvents</span></span><br><span data-ttu-id="511f6-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="511f6-648">
         - File</span></span><br><span data-ttu-id="511f6-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-649">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="511f6-650">
         -PdfFile</span></span><br><span data-ttu-id="511f6-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-651">
         - Selection</span></span><br><span data-ttu-id="511f6-652">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-652">
         - Settings</span></span><br><span data-ttu-id="511f6-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-653">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="511f6-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="511f6-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="511f6-655">Платформа</span><span class="sxs-lookup"><span data-stu-id="511f6-655">Platform</span></span></th>
    <th><span data-ttu-id="511f6-656">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="511f6-656">Extension points</span></span></th>
    <th><span data-ttu-id="511f6-657">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="511f6-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="511f6-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="511f6-659">Office Online</span></span></td>
    <td> <span data-ttu-id="511f6-660">- Контент</span><span class="sxs-lookup"><span data-stu-id="511f6-660">- Content</span></span><br><span data-ttu-id="511f6-661">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-661">
         - Taskpane</span></span><br><span data-ttu-id="511f6-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="511f6-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="511f6-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="511f6-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="511f6-665">-DocumentEvents</span></span><br><span data-ttu-id="511f6-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-666">
         -HtmlCoercion</span></span><br><span data-ttu-id="511f6-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-667">
         -ImageCoercion</span></span><br><span data-ttu-id="511f6-668">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="511f6-668">
         - Settings</span></span><br><span data-ttu-id="511f6-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-669">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="511f6-670">Project</span><span class="sxs-lookup"><span data-stu-id="511f6-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="511f6-671">Платформа</span><span class="sxs-lookup"><span data-stu-id="511f6-671">Platform</span></span></th>
    <th><span data-ttu-id="511f6-672">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="511f6-672">Extension points</span></span></th>
    <th><span data-ttu-id="511f6-673">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="511f6-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные API</b></a></span><span class="sxs-lookup"><span data-stu-id="511f6-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-675">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-676">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-676">- Taskpane</span></span></td>
    <td> <span data-ttu-id="511f6-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-678">- Selection</span></span><br><span data-ttu-id="511f6-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-679">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-680">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="511f6-681">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-681">- Taskpane</span></span></td>
    <td> <span data-ttu-id="511f6-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-683">- Selection</span></span><br><span data-ttu-id="511f6-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-684">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="511f6-685">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="511f6-685">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="511f6-686">- Область задач</span><span class="sxs-lookup"><span data-stu-id="511f6-686">- Taskpane</span></span></td>
    <td> <span data-ttu-id="511f6-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="511f6-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="511f6-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="511f6-688">- Selection</span></span><br><span data-ttu-id="511f6-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="511f6-689">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="511f6-690">См. также</span><span class="sxs-lookup"><span data-stu-id="511f6-690">See also</span></span>

- [<span data-ttu-id="511f6-691">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="511f6-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="511f6-692">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="511f6-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="511f6-693">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="511f6-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="511f6-694">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="511f6-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
