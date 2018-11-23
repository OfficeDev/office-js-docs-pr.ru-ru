---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 11/07/2018
ms.openlocfilehash: c3da40be21c0e569028dd10e93e33760ba2bd39d
ms.sourcegitcommit: 3e84d616e69f39eeeeea773f2431e7d674c4a9f5
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/22/2018
ms.locfileid: "26644755"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7743f-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7743f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7743f-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="7743f-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="7743f-105">В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="7743f-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="7743f-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="7743f-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="7743f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="7743f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7743f-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="7743f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7743f-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="7743f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7743f-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7743f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="7743f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="7743f-113">Office Online</span></span></td>
    <td> <span data-ttu-id="7743f-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-114">- TaskPane</span></span><br><span data-ttu-id="7743f-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-115">
        - Content</span></span><br><span data-ttu-id="7743f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="7743f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7743f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7743f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7743f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7743f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7743f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7743f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7743f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7743f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7743f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7743f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-126">
        - BindingEvents</span></span><br><span data-ttu-id="7743f-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-127">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-128">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-129">
        - File</span></span><br><span data-ttu-id="7743f-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-130">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-132">
        - Selection</span></span><br><span data-ttu-id="7743f-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-133">
        - Settings</span></span><br><span data-ttu-id="7743f-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-134">
        - TableBindings</span></span><br><span data-ttu-id="7743f-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-135">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-136">
        - TextBindings</span></span><br><span data-ttu-id="7743f-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="7743f-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-139">
        - TaskPane</span></span><br><span data-ttu-id="7743f-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7743f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-142">
        - BindingEvents</span></span><br><span data-ttu-id="7743f-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-143">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-144">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-145">
        - File</span></span><br><span data-ttu-id="7743f-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-146">
        - ImageCoercion</span></span><br><span data-ttu-id="7743f-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-147">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-149">
        - Selection</span></span><br><span data-ttu-id="7743f-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-150">
        - Settings</span></span><br><span data-ttu-id="7743f-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-151">
        - TableBindings</span></span><br><span data-ttu-id="7743f-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-152">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-153">
        - TextBindings</span></span><br><span data-ttu-id="7743f-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="7743f-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-156">- TaskPane</span></span><br><span data-ttu-id="7743f-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-157">
        - Content</span></span><br><span data-ttu-id="7743f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7743f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7743f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7743f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7743f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7743f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7743f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7743f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7743f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7743f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7743f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-168">- BindingEvents</span></span><br><span data-ttu-id="7743f-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-169">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-170">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-171">
        - File</span></span><br><span data-ttu-id="7743f-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-172">
        - ImageCoercion</span></span><br><span data-ttu-id="7743f-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-173">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-175">
        - Selection</span></span><br><span data-ttu-id="7743f-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-176">
        - Settings</span></span><br><span data-ttu-id="7743f-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-177">
        - TableBindings</span></span><br><span data-ttu-id="7743f-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-178">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-179">
        - TextBindings</span></span><br><span data-ttu-id="7743f-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-181">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="7743f-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-182">- TaskPane</span></span><br><span data-ttu-id="7743f-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-183">
        - Content</span></span><br><span data-ttu-id="7743f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7743f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7743f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7743f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7743f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7743f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7743f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7743f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7743f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7743f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7743f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-194">- BindingEvents</span></span><br><span data-ttu-id="7743f-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-195">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-196">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-197">
        - File</span></span><br><span data-ttu-id="7743f-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-198">
        - ImageCoercion</span></span><br><span data-ttu-id="7743f-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-199">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-201">
        - Selection</span></span><br><span data-ttu-id="7743f-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-202">
        - Settings</span></span><br><span data-ttu-id="7743f-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-203">
        - TableBindings</span></span><br><span data-ttu-id="7743f-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-204">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-205">
        - TextBindings</span></span><br><span data-ttu-id="7743f-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-207">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="7743f-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="7743f-208">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-208">- TaskPane</span></span><br><span data-ttu-id="7743f-209">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-209">
        - Content</span></span></td>
    <td><span data-ttu-id="7743f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7743f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7743f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7743f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7743f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7743f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7743f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7743f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7743f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7743f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-219">- BindingEvents</span></span><br><span data-ttu-id="7743f-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-220">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-221">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-222">
        - File</span></span><br><span data-ttu-id="7743f-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-223">
        - ImageCoercion</span></span><br><span data-ttu-id="7743f-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-224">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-226">
        - Selection</span></span><br><span data-ttu-id="7743f-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-227">
        - Settings</span></span><br><span data-ttu-id="7743f-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-228">
        - TableBindings</span></span><br><span data-ttu-id="7743f-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-229">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-230">
        - TextBindings</span></span><br><span data-ttu-id="7743f-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-232">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="7743f-233">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-233">- TaskPane</span></span><br><span data-ttu-id="7743f-234">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-234">
        - Content</span></span><br><span data-ttu-id="7743f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7743f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7743f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7743f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7743f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7743f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7743f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7743f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7743f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7743f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7743f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-245">- BindingEvents</span></span><br><span data-ttu-id="7743f-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-246">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-247">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-248">
        - File</span></span><br><span data-ttu-id="7743f-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-249">
        - ImageCoercion</span></span><br><span data-ttu-id="7743f-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-250">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-252">
        - PdfFile</span></span><br><span data-ttu-id="7743f-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-253">
        - Selection</span></span><br><span data-ttu-id="7743f-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-254">
        - Settings</span></span><br><span data-ttu-id="7743f-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-255">
        - TableBindings</span></span><br><span data-ttu-id="7743f-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-256">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-257">
        - TextBindings</span></span><br><span data-ttu-id="7743f-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-259">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="7743f-260">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-260">- TaskPane</span></span><br><span data-ttu-id="7743f-261">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-261">
        - Content</span></span><br><span data-ttu-id="7743f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7743f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7743f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7743f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7743f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7743f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7743f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7743f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7743f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7743f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7743f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7743f-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-272">- BindingEvents</span></span><br><span data-ttu-id="7743f-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-273">
        - CompressedFile</span></span><br><span data-ttu-id="7743f-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-274">
        - DocumentEvents</span></span><br><span data-ttu-id="7743f-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="7743f-275">
        - File</span></span><br><span data-ttu-id="7743f-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-276">
        - ImageCoercion</span></span><br><span data-ttu-id="7743f-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-277">
        - MatrixBindings</span></span><br><span data-ttu-id="7743f-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="7743f-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-279">
        - PdfFile</span></span><br><span data-ttu-id="7743f-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-280">
        - Selection</span></span><br><span data-ttu-id="7743f-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-281">
        - Settings</span></span><br><span data-ttu-id="7743f-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-282">
        - TableBindings</span></span><br><span data-ttu-id="7743f-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-283">
        - TableCoercion</span></span><br><span data-ttu-id="7743f-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-284">
        - TextBindings</span></span><br><span data-ttu-id="7743f-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="7743f-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="7743f-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7743f-287">Платформа</span><span class="sxs-lookup"><span data-stu-id="7743f-287">Platform</span></span></th>
    <th><span data-ttu-id="7743f-288">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="7743f-288">Extension points</span></span></th>
    <th><span data-ttu-id="7743f-289">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="7743f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="7743f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="7743f-291">Office Online</span></span></td>
    <td> <span data-ttu-id="7743f-292">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-292">- Mail Read</span></span><br><span data-ttu-id="7743f-293">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="7743f-293">
      - Mail Compose</span></span><br><span data-ttu-id="7743f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7743f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7743f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7743f-302">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-303">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-304">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-304">- Mail Read</span></span><br><span data-ttu-id="7743f-305">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="7743f-305">
      - Mail Compose</span></span><br><span data-ttu-id="7743f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="7743f-311">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-312">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-313">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-313">- Mail Read</span></span><br><span data-ttu-id="7743f-314">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="7743f-314">
      - Mail Compose</span></span><br><span data-ttu-id="7743f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7743f-316">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="7743f-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7743f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7743f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7743f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7743f-324">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-325">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-326">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-326">- Mail Read</span></span><br><span data-ttu-id="7743f-327">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="7743f-327">
      - Mail Compose</span></span><br><span data-ttu-id="7743f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7743f-329">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="7743f-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7743f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7743f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7743f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7743f-337">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-338">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="7743f-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="7743f-339">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-339">- Mail Read</span></span><br><span data-ttu-id="7743f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7743f-346">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-347">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7743f-348">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-348">- Mail Read</span></span><br><span data-ttu-id="7743f-349">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="7743f-349">
      - Mail Compose</span></span><br><span data-ttu-id="7743f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7743f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7743f-357">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-358">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="7743f-359">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-359">- Mail Read</span></span><br><span data-ttu-id="7743f-360">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="7743f-360">
      - Mail Compose</span></span><br><span data-ttu-id="7743f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7743f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7743f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7743f-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7743f-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7743f-369">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-370">Office для Android</span><span class="sxs-lookup"><span data-stu-id="7743f-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="7743f-371">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="7743f-371">- Mail Read</span></span><br><span data-ttu-id="7743f-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7743f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7743f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7743f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7743f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7743f-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7743f-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7743f-378">Недоступно</span><span class="sxs-lookup"><span data-stu-id="7743f-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="7743f-379">Word</span><span class="sxs-lookup"><span data-stu-id="7743f-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7743f-380">Платформа</span><span class="sxs-lookup"><span data-stu-id="7743f-380">Platform</span></span></th>
    <th><span data-ttu-id="7743f-381">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="7743f-381">Extension points</span></span></th>
    <th><span data-ttu-id="7743f-382">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="7743f-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="7743f-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="7743f-384">Office Online</span></span></td>
    <td> <span data-ttu-id="7743f-385">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-385">- TaskPane</span></span><br><span data-ttu-id="7743f-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7743f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7743f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7743f-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-391">- BindingEvents</span></span><br><span data-ttu-id="7743f-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-393">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-394">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-394">
         - File</span></span><br><span data-ttu-id="7743f-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-396">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-397">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-400">
         - PdfFile</span></span><br><span data-ttu-id="7743f-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-401">
         - Selection</span></span><br><span data-ttu-id="7743f-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-402">
         - Settings</span></span><br><span data-ttu-id="7743f-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-403">
         - TableBindings</span></span><br><span data-ttu-id="7743f-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-404">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-405">
         - TextBindings</span></span><br><span data-ttu-id="7743f-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-406">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-408">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-409">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7743f-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-411">- BindingEvents</span></span><br><span data-ttu-id="7743f-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-412">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-414">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-415">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-415">
         - File</span></span><br><span data-ttu-id="7743f-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-417">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-418">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-421">
         - PdfFile</span></span><br><span data-ttu-id="7743f-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-422">
         - Selection</span></span><br><span data-ttu-id="7743f-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-423">
         - Settings</span></span><br><span data-ttu-id="7743f-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-424">
         - TableBindings</span></span><br><span data-ttu-id="7743f-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-425">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-426">
         - TextBindings</span></span><br><span data-ttu-id="7743f-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-427">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-429">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-430">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-430">- TaskPane</span></span><br><span data-ttu-id="7743f-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7743f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7743f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7743f-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-436">- BindingEvents</span></span><br><span data-ttu-id="7743f-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-437">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-439">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-440">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-440">
         - File</span></span><br><span data-ttu-id="7743f-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-442">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-443">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-446">
         - PdfFile</span></span><br><span data-ttu-id="7743f-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-447">
         - Selection</span></span><br><span data-ttu-id="7743f-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-448">
         - Settings</span></span><br><span data-ttu-id="7743f-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-449">
         - TableBindings</span></span><br><span data-ttu-id="7743f-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-450">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-451">
         - TextBindings</span></span><br><span data-ttu-id="7743f-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-452">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-454">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-455">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-455">- TaskPane</span></span><br><span data-ttu-id="7743f-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7743f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7743f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7743f-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-461">- BindingEvents</span></span><br><span data-ttu-id="7743f-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-462">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-464">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-465">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-465">
         - File</span></span><br><span data-ttu-id="7743f-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-467">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-468">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-471">
         - PdfFile</span></span><br><span data-ttu-id="7743f-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-472">
         - Selection</span></span><br><span data-ttu-id="7743f-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-473">
         - Settings</span></span><br><span data-ttu-id="7743f-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-474">
         - TableBindings</span></span><br><span data-ttu-id="7743f-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-475">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-476">
         - TextBindings</span></span><br><span data-ttu-id="7743f-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-477">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-479">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="7743f-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="7743f-480">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7743f-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7743f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7743f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7743f-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7743f-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7743f-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-485">- BindingEvents</span></span><br><span data-ttu-id="7743f-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-486">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-488">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-489">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-489">
         - File</span></span><br><span data-ttu-id="7743f-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-491">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-492">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-495">
         - PdfFile</span></span><br><span data-ttu-id="7743f-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-496">
         - Selection</span></span><br><span data-ttu-id="7743f-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-497">
         - Settings</span></span><br><span data-ttu-id="7743f-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-498">
         - TableBindings</span></span><br><span data-ttu-id="7743f-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-499">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-500">
         - TextBindings</span></span><br><span data-ttu-id="7743f-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-501">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-503">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7743f-504">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-504">- TaskPane</span></span><br><span data-ttu-id="7743f-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7743f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7743f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7743f-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7743f-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7743f-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-510">- BindingEvents</span></span><br><span data-ttu-id="7743f-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-511">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-513">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-514">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-514">
         - File</span></span><br><span data-ttu-id="7743f-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-516">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-517">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-520">
         - PdfFile</span></span><br><span data-ttu-id="7743f-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-521">
         - Selection</span></span><br><span data-ttu-id="7743f-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-522">
         - Settings</span></span><br><span data-ttu-id="7743f-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-523">
         - TableBindings</span></span><br><span data-ttu-id="7743f-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-524">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-525">
         - TextBindings</span></span><br><span data-ttu-id="7743f-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-526">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-528">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="7743f-529">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-529">- TaskPane</span></span><br><span data-ttu-id="7743f-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7743f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7743f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7743f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7743f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7743f-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7743f-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7743f-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-535">- BindingEvents</span></span><br><span data-ttu-id="7743f-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-536">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7743f-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="7743f-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-538">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="7743f-539">
         - File</span></span><br><span data-ttu-id="7743f-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-541">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-542">
         - MatrixBindings</span></span><br><span data-ttu-id="7743f-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="7743f-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7743f-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-545">
         - PdfFile</span></span><br><span data-ttu-id="7743f-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-546">
         - Selection</span></span><br><span data-ttu-id="7743f-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7743f-547">
         - Settings</span></span><br><span data-ttu-id="7743f-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-548">
         - TableBindings</span></span><br><span data-ttu-id="7743f-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-549">
         - TableCoercion</span></span><br><span data-ttu-id="7743f-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7743f-550">
         - TextBindings</span></span><br><span data-ttu-id="7743f-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-551">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7743f-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7743f-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7743f-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7743f-554">Платформа</span><span class="sxs-lookup"><span data-stu-id="7743f-554">Platform</span></span></th>
    <th><span data-ttu-id="7743f-555">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="7743f-555">Extension points</span></span></th>
    <th><span data-ttu-id="7743f-556">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="7743f-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="7743f-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="7743f-558">Office Online</span></span></td>
    <td> <span data-ttu-id="7743f-559">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-559">- Content</span></span><br><span data-ttu-id="7743f-560">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-560">
         - TaskPane</span></span><br><span data-ttu-id="7743f-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-563">- ActiveView</span></span><br><span data-ttu-id="7743f-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-564">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-565">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-566">
         - File</span></span><br><span data-ttu-id="7743f-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-567">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-568">
         - PdfFile</span></span><br><span data-ttu-id="7743f-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-569">
         - Selection</span></span><br><span data-ttu-id="7743f-570">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-570">
         - Settings</span></span><br><span data-ttu-id="7743f-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-572">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-573">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-573">- Content</span></span><br><span data-ttu-id="7743f-574">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="7743f-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7743f-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7743f-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-576">- ActiveView</span></span><br><span data-ttu-id="7743f-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-577">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-578">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-579">
         - File</span></span><br><span data-ttu-id="7743f-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-580">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-581">
         - PdfFile</span></span><br><span data-ttu-id="7743f-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-582">
         - Selection</span></span><br><span data-ttu-id="7743f-583">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-583">
         - Settings</span></span><br><span data-ttu-id="7743f-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-586">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-586">- Content</span></span><br><span data-ttu-id="7743f-587">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-587">
         - TaskPane</span></span><br><span data-ttu-id="7743f-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-590">- ActiveView</span></span><br><span data-ttu-id="7743f-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-591">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-592">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-593">
         - File</span></span><br><span data-ttu-id="7743f-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-594">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-595">
         - PdfFile</span></span><br><span data-ttu-id="7743f-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-596">
         - Selection</span></span><br><span data-ttu-id="7743f-597">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-597">
         - Settings</span></span><br><span data-ttu-id="7743f-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-599">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-600">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-600">- Content</span></span><br><span data-ttu-id="7743f-601">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-601">
         - TaskPane</span></span><br><span data-ttu-id="7743f-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-604">- ActiveView</span></span><br><span data-ttu-id="7743f-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-605">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-606">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-607">
         - File</span></span><br><span data-ttu-id="7743f-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-608">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-609">
         - PdfFile</span></span><br><span data-ttu-id="7743f-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-610">
         - Selection</span></span><br><span data-ttu-id="7743f-611">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-611">
         - Settings</span></span><br><span data-ttu-id="7743f-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-613">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="7743f-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="7743f-614">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-614">- Content</span></span><br><span data-ttu-id="7743f-615">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7743f-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="7743f-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-617">- ActiveView</span></span><br><span data-ttu-id="7743f-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-618">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-619">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-620">
         - File</span></span><br><span data-ttu-id="7743f-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-621">
         - PdfFile</span></span><br><span data-ttu-id="7743f-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-622">
         - Selection</span></span><br><span data-ttu-id="7743f-623">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-623">
         - Settings</span></span><br><span data-ttu-id="7743f-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-624">
         - TextCoercion</span></span><br><span data-ttu-id="7743f-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-626">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7743f-627">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-627">- Content</span></span><br><span data-ttu-id="7743f-628">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-628">
         - TaskPane</span></span><br><span data-ttu-id="7743f-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-631">- ActiveView</span></span><br><span data-ttu-id="7743f-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-632">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-633">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-634">
         - File</span></span><br><span data-ttu-id="7743f-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-635">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-636">
         - PdfFile</span></span><br><span data-ttu-id="7743f-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-637">
         - Selection</span></span><br><span data-ttu-id="7743f-638">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-638">
         - Settings</span></span><br><span data-ttu-id="7743f-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-640">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="7743f-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="7743f-641">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-641">- Content</span></span><br><span data-ttu-id="7743f-642">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-642">
         - TaskPane</span></span><br><span data-ttu-id="7743f-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7743f-645">- ActiveView</span></span><br><span data-ttu-id="7743f-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7743f-646">
         - CompressedFile</span></span><br><span data-ttu-id="7743f-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-647">
         - DocumentEvents</span></span><br><span data-ttu-id="7743f-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="7743f-648">
         - File</span></span><br><span data-ttu-id="7743f-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-649">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7743f-650">
         - PdfFile</span></span><br><span data-ttu-id="7743f-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-651">
         - Selection</span></span><br><span data-ttu-id="7743f-652">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-652">
         - Settings</span></span><br><span data-ttu-id="7743f-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="7743f-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="7743f-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7743f-655">Платформа</span><span class="sxs-lookup"><span data-stu-id="7743f-655">Platform</span></span></th>
    <th><span data-ttu-id="7743f-656">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="7743f-656">Extension points</span></span></th>
    <th><span data-ttu-id="7743f-657">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="7743f-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="7743f-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="7743f-659">Office Online</span></span></td>
    <td> <span data-ttu-id="7743f-660">- Контент</span><span class="sxs-lookup"><span data-stu-id="7743f-660">- Content</span></span><br><span data-ttu-id="7743f-661">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-661">
         - TaskPane</span></span><br><span data-ttu-id="7743f-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="7743f-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7743f-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7743f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7743f-665">- DocumentEvents</span></span><br><span data-ttu-id="7743f-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="7743f-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-667">
         - ImageCoercion</span></span><br><span data-ttu-id="7743f-668">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="7743f-668">
         - Settings</span></span><br><span data-ttu-id="7743f-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="7743f-670">Project</span><span class="sxs-lookup"><span data-stu-id="7743f-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7743f-671">Платформа</span><span class="sxs-lookup"><span data-stu-id="7743f-671">Platform</span></span></th>
    <th><span data-ttu-id="7743f-672">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="7743f-672">Extension points</span></span></th>
    <th><span data-ttu-id="7743f-673">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="7743f-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные API</b></a></span><span class="sxs-lookup"><span data-stu-id="7743f-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-675">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-676">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7743f-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-678">- Selection</span></span><br><span data-ttu-id="7743f-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-680">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-681">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7743f-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-683">- Selection</span></span><br><span data-ttu-id="7743f-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7743f-685">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="7743f-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="7743f-686">- Область задач</span><span class="sxs-lookup"><span data-stu-id="7743f-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7743f-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7743f-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7743f-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="7743f-688">- Selection</span></span><br><span data-ttu-id="7743f-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7743f-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="7743f-690">См. также</span><span class="sxs-lookup"><span data-stu-id="7743f-690">See also</span></span>

- [<span data-ttu-id="7743f-691">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7743f-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7743f-692">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="7743f-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="7743f-693">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="7743f-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="7743f-694">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="7743f-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
