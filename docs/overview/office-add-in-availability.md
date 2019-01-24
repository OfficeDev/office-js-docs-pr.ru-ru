---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 11/07/2018
localization_priority: Priority
ms.openlocfilehash: 9f8b94483d22f24dcb0a6a2ad99df6167533133f
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388341"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5aa2c-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5aa2c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5aa2c-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="5aa2c-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="5aa2c-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="5aa2c-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="5aa2c-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="5aa2c-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="5aa2c-108">Excel</span><span class="sxs-lookup"><span data-stu-id="5aa2c-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5aa2c-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="5aa2c-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5aa2c-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5aa2c-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5aa2c-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5aa2c-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="5aa2c-113">Office Online</span></span></td>
    <td> <span data-ttu-id="5aa2c-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-114">- TaskPane</span></span><br><span data-ttu-id="5aa2c-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-115">
        - Content</span></span><br><span data-ttu-id="5aa2c-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="5aa2c-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5aa2c-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5aa2c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5aa2c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5aa2c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5aa2c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5aa2c-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-126">
        - BindingEvents</span></span><br><span data-ttu-id="5aa2c-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-127">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-128">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-129">
        - File</span></span><br><span data-ttu-id="5aa2c-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-130">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-132">
        - Selection</span></span><br><span data-ttu-id="5aa2c-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-133">
        - Settings</span></span><br><span data-ttu-id="5aa2c-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-134">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-135">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-136">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="5aa2c-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-139">
        - TaskPane</span></span><br><span data-ttu-id="5aa2c-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5aa2c-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-142">
        - BindingEvents</span></span><br><span data-ttu-id="5aa2c-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-143">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-144">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-145">
        - File</span></span><br><span data-ttu-id="5aa2c-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-146">
        - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-147">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-149">
        - Selection</span></span><br><span data-ttu-id="5aa2c-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-150">
        - Settings</span></span><br><span data-ttu-id="5aa2c-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-151">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-152">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-153">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="5aa2c-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-156">- TaskPane</span></span><br><span data-ttu-id="5aa2c-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-157">
        - Content</span></span><br><span data-ttu-id="5aa2c-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5aa2c-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5aa2c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5aa2c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5aa2c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5aa2c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5aa2c-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-168">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-169">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-170">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-171">
        - File</span></span><br><span data-ttu-id="5aa2c-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-172">
        - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-173">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-175">
        - Selection</span></span><br><span data-ttu-id="5aa2c-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-176">
        - Settings</span></span><br><span data-ttu-id="5aa2c-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-177">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-178">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-179">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-181">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="5aa2c-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-182">- TaskPane</span></span><br><span data-ttu-id="5aa2c-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-183">
        - Content</span></span><br><span data-ttu-id="5aa2c-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5aa2c-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5aa2c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5aa2c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5aa2c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5aa2c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5aa2c-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-194">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-195">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-196">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-197">
        - File</span></span><br><span data-ttu-id="5aa2c-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-198">
        - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-199">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-201">
        - Selection</span></span><br><span data-ttu-id="5aa2c-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-202">
        - Settings</span></span><br><span data-ttu-id="5aa2c-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-203">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-204">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-205">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-207">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="5aa2c-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="5aa2c-208">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-208">- TaskPane</span></span><br><span data-ttu-id="5aa2c-209">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-209">
        - Content</span></span></td>
    <td><span data-ttu-id="5aa2c-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5aa2c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5aa2c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5aa2c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5aa2c-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5aa2c-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-219">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-220">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-221">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-222">
        - File</span></span><br><span data-ttu-id="5aa2c-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-223">
        - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-224">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-226">
        - Selection</span></span><br><span data-ttu-id="5aa2c-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-227">
        - Settings</span></span><br><span data-ttu-id="5aa2c-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-228">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-229">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-230">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-232">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="5aa2c-233">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-233">- TaskPane</span></span><br><span data-ttu-id="5aa2c-234">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-234">
        - Content</span></span><br><span data-ttu-id="5aa2c-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5aa2c-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5aa2c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5aa2c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5aa2c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5aa2c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5aa2c-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-245">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-246">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-247">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-248">
        - File</span></span><br><span data-ttu-id="5aa2c-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-249">
        - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-250">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-252">
        - PdfFile</span></span><br><span data-ttu-id="5aa2c-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-253">
        - Selection</span></span><br><span data-ttu-id="5aa2c-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-254">
        - Settings</span></span><br><span data-ttu-id="5aa2c-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-255">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-256">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-257">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-259">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="5aa2c-260">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-260">- TaskPane</span></span><br><span data-ttu-id="5aa2c-261">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-261">
        - Content</span></span><br><span data-ttu-id="5aa2c-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5aa2c-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5aa2c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5aa2c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5aa2c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5aa2c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5aa2c-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5aa2c-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-272">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-273">
        - CompressedFile</span></span><br><span data-ttu-id="5aa2c-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-274">
        - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-275">
        - File</span></span><br><span data-ttu-id="5aa2c-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-276">
        - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-277">
        - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-279">
        - PdfFile</span></span><br><span data-ttu-id="5aa2c-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-280">
        - Selection</span></span><br><span data-ttu-id="5aa2c-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-281">
        - Settings</span></span><br><span data-ttu-id="5aa2c-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-282">
        - TableBindings</span></span><br><span data-ttu-id="5aa2c-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-283">
        - TableCoercion</span></span><br><span data-ttu-id="5aa2c-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-284">
        - TextBindings</span></span><br><span data-ttu-id="5aa2c-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="5aa2c-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="5aa2c-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5aa2c-287">Платформа</span><span class="sxs-lookup"><span data-stu-id="5aa2c-287">Platform</span></span></th>
    <th><span data-ttu-id="5aa2c-288">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5aa2c-288">Extension points</span></span></th>
    <th><span data-ttu-id="5aa2c-289">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="5aa2c-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="5aa2c-291">Office Online</span></span></td>
    <td> <span data-ttu-id="5aa2c-292">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-292">- Mail Read</span></span><br><span data-ttu-id="5aa2c-293">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-293">
      - Mail Compose</span></span><br><span data-ttu-id="5aa2c-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5aa2c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5aa2c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5aa2c-302">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-303">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-304">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-304">- Mail Read</span></span><br><span data-ttu-id="5aa2c-305">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-305">
      - Mail Compose</span></span><br><span data-ttu-id="5aa2c-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="5aa2c-311">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-312">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-313">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-313">- Mail Read</span></span><br><span data-ttu-id="5aa2c-314">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-314">
      - Mail Compose</span></span><br><span data-ttu-id="5aa2c-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5aa2c-316">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="5aa2c-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5aa2c-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5aa2c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5aa2c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5aa2c-324">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-325">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-326">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-326">- Mail Read</span></span><br><span data-ttu-id="5aa2c-327">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-327">
      - Mail Compose</span></span><br><span data-ttu-id="5aa2c-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5aa2c-329">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="5aa2c-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5aa2c-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5aa2c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5aa2c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5aa2c-337">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-338">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="5aa2c-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5aa2c-339">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-339">- Mail Read</span></span><br><span data-ttu-id="5aa2c-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5aa2c-346">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-347">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5aa2c-348">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-348">- Mail Read</span></span><br><span data-ttu-id="5aa2c-349">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-349">
      - Mail Compose</span></span><br><span data-ttu-id="5aa2c-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5aa2c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5aa2c-357">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-358">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="5aa2c-359">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-359">- Mail Read</span></span><br><span data-ttu-id="5aa2c-360">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-360">
      - Mail Compose</span></span><br><span data-ttu-id="5aa2c-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5aa2c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5aa2c-368">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-368">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-369">Office для Android</span><span class="sxs-lookup"><span data-stu-id="5aa2c-369">Office for Android</span></span></td>
    <td> <span data-ttu-id="5aa2c-370">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="5aa2c-370">- Mail Read</span></span><br><span data-ttu-id="5aa2c-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5aa2c-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5aa2c-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5aa2c-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5aa2c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5aa2c-377">Недоступно</span><span class="sxs-lookup"><span data-stu-id="5aa2c-377">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="5aa2c-378">Word</span><span class="sxs-lookup"><span data-stu-id="5aa2c-378">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5aa2c-379">Платформа</span><span class="sxs-lookup"><span data-stu-id="5aa2c-379">Platform</span></span></th>
    <th><span data-ttu-id="5aa2c-380">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5aa2c-380">Extension points</span></span></th>
    <th><span data-ttu-id="5aa2c-381">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-381">API requirement sets</span></span></th>
    <th><span data-ttu-id="5aa2c-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-383">Office Online</span><span class="sxs-lookup"><span data-stu-id="5aa2c-383">Office Online</span></span></td>
    <td> <span data-ttu-id="5aa2c-384">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-384">- TaskPane</span></span><br><span data-ttu-id="5aa2c-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-390">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-390">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-391">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-391">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-392">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-392">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-393">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-393">
         - File</span></span><br><span data-ttu-id="5aa2c-394">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-394">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-395">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-395">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-396">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-396">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-397">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-397">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-398">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-398">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-399">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-399">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-400">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-400">
         - Selection</span></span><br><span data-ttu-id="5aa2c-401">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-401">
         - Settings</span></span><br><span data-ttu-id="5aa2c-402">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-402">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-403">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-403">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-404">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-404">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-405">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-405">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-406">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-406">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-407">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-407">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-408">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-408">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5aa2c-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-410">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-410">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-411">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-411">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-412">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-412">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-413">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-413">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-414">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-414">
         - File</span></span><br><span data-ttu-id="5aa2c-415">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-415">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-416">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-416">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-417">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-417">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-418">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-418">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-419">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-419">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-420">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-420">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-421">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-421">
         - Selection</span></span><br><span data-ttu-id="5aa2c-422">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-422">
         - Settings</span></span><br><span data-ttu-id="5aa2c-423">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-423">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-424">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-424">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-425">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-425">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-426">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-426">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-427">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-427">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-428">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-428">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-429">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-429">- TaskPane</span></span><br><span data-ttu-id="5aa2c-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-435">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-435">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-436">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-436">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-437">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-437">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-438">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-438">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-439">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-439">
         - File</span></span><br><span data-ttu-id="5aa2c-440">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-440">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-441">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-441">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-442">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-442">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-443">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-443">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-444">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-444">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-445">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-445">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-446">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-446">
         - Selection</span></span><br><span data-ttu-id="5aa2c-447">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-447">
         - Settings</span></span><br><span data-ttu-id="5aa2c-448">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-448">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-449">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-449">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-450">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-450">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-451">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-451">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-452">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-452">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-453">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-453">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-454">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-454">- TaskPane</span></span><br><span data-ttu-id="5aa2c-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-460">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-460">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-461">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-461">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-462">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-462">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-463">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-463">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-464">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-464">
         - File</span></span><br><span data-ttu-id="5aa2c-465">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-465">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-466">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-466">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-467">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-467">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-468">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-468">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-469">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-469">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-470">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-470">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-471">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-471">
         - Selection</span></span><br><span data-ttu-id="5aa2c-472">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-472">
         - Settings</span></span><br><span data-ttu-id="5aa2c-473">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-473">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-474">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-474">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-475">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-475">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-476">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-476">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-477">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-477">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-478">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="5aa2c-478">Office for iPad</span></span></td>
    <td> <span data-ttu-id="5aa2c-479">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-479">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5aa2c-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5aa2c-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5aa2c-484">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-484">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-485">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-485">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-486">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-486">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-487">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-488">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-488">
         - File</span></span><br><span data-ttu-id="5aa2c-489">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-489">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-490">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-491">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-491">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-492">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-492">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-493">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-493">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-494">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-494">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-495">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-495">
         - Selection</span></span><br><span data-ttu-id="5aa2c-496">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-496">
         - Settings</span></span><br><span data-ttu-id="5aa2c-497">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-497">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-498">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-498">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-499">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-499">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-500">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-500">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-501">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-501">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-502">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-502">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5aa2c-503">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-503">- TaskPane</span></span><br><span data-ttu-id="5aa2c-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5aa2c-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5aa2c-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-509">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-510">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-510">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-511">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-511">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-512">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-512">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-513">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-513">
         - File</span></span><br><span data-ttu-id="5aa2c-514">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-514">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-515">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-515">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-516">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-519">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-520">
         - Selection</span></span><br><span data-ttu-id="5aa2c-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-521">
         - Settings</span></span><br><span data-ttu-id="5aa2c-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-522">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-523">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-524">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-525">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-526">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-527">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-527">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="5aa2c-528">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-528">- TaskPane</span></span><br><span data-ttu-id="5aa2c-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5aa2c-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5aa2c-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5aa2c-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5aa2c-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-534">- BindingEvents</span></span><br><span data-ttu-id="5aa2c-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-535">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5aa2c-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="5aa2c-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-537">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-538">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="5aa2c-538">
         - File</span></span><br><span data-ttu-id="5aa2c-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-540">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-540">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-541">
         - MatrixBindings</span></span><br><span data-ttu-id="5aa2c-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="5aa2c-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5aa2c-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-544">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-545">
         - Selection</span></span><br><span data-ttu-id="5aa2c-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-546">
         - Settings</span></span><br><span data-ttu-id="5aa2c-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-547">
         - TableBindings</span></span><br><span data-ttu-id="5aa2c-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-548">
         - TableCoercion</span></span><br><span data-ttu-id="5aa2c-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5aa2c-549">
         - TextBindings</span></span><br><span data-ttu-id="5aa2c-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-550">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-551">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5aa2c-552">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5aa2c-552">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5aa2c-553">Платформа</span><span class="sxs-lookup"><span data-stu-id="5aa2c-553">Platform</span></span></th>
    <th><span data-ttu-id="5aa2c-554">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5aa2c-554">Extension points</span></span></th>
    <th><span data-ttu-id="5aa2c-555">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-555">API requirement sets</span></span></th>
    <th><span data-ttu-id="5aa2c-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-557">Office Online</span><span class="sxs-lookup"><span data-stu-id="5aa2c-557">Office Online</span></span></td>
    <td> <span data-ttu-id="5aa2c-558">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-558">- Content</span></span><br><span data-ttu-id="5aa2c-559">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-559">
         - TaskPane</span></span><br><span data-ttu-id="5aa2c-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-562">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-562">- ActiveView</span></span><br><span data-ttu-id="5aa2c-563">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-563">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-564">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-565">
         - File</span></span><br><span data-ttu-id="5aa2c-566">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-566">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-567">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-568">
         - Selection</span></span><br><span data-ttu-id="5aa2c-569">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-569">
         - Settings</span></span><br><span data-ttu-id="5aa2c-570">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-570">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-571">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-571">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-572">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-572">- Content</span></span><br><span data-ttu-id="5aa2c-573">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-573">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="5aa2c-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5aa2c-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5aa2c-575">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-575">- ActiveView</span></span><br><span data-ttu-id="5aa2c-576">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-576">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-577">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-577">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-578">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-578">
         - File</span></span><br><span data-ttu-id="5aa2c-579">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-579">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-580">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-580">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-581">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-581">
         - Selection</span></span><br><span data-ttu-id="5aa2c-582">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-582">
         - Settings</span></span><br><span data-ttu-id="5aa2c-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-583">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-584">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-584">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-585">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-585">- Content</span></span><br><span data-ttu-id="5aa2c-586">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-586">
         - TaskPane</span></span><br><span data-ttu-id="5aa2c-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-589">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-589">- ActiveView</span></span><br><span data-ttu-id="5aa2c-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-590">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-591">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-592">
         - File</span></span><br><span data-ttu-id="5aa2c-593">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-593">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-594">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-594">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-595">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-595">
         - Selection</span></span><br><span data-ttu-id="5aa2c-596">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-596">
         - Settings</span></span><br><span data-ttu-id="5aa2c-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-597">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-598">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-598">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-599">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-599">- Content</span></span><br><span data-ttu-id="5aa2c-600">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-600">
         - TaskPane</span></span><br><span data-ttu-id="5aa2c-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-603">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-603">- ActiveView</span></span><br><span data-ttu-id="5aa2c-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-604">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-605">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-605">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-606">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-606">
         - File</span></span><br><span data-ttu-id="5aa2c-607">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-607">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-608">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-609">
         - Selection</span></span><br><span data-ttu-id="5aa2c-610">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-610">
         - Settings</span></span><br><span data-ttu-id="5aa2c-611">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-611">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-612">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="5aa2c-612">Office for iPad</span></span></td>
    <td> <span data-ttu-id="5aa2c-613">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-613">- Content</span></span><br><span data-ttu-id="5aa2c-614">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-614">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5aa2c-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="5aa2c-616">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-616">- ActiveView</span></span><br><span data-ttu-id="5aa2c-617">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-617">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-618">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-619">
         - File</span></span><br><span data-ttu-id="5aa2c-620">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-620">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-621">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-621">
         - Selection</span></span><br><span data-ttu-id="5aa2c-622">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-622">
         - Settings</span></span><br><span data-ttu-id="5aa2c-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-623">
         - TextCoercion</span></span><br><span data-ttu-id="5aa2c-624">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-624">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-625">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-625">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5aa2c-626">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-626">- Content</span></span><br><span data-ttu-id="5aa2c-627">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-627">
         - TaskPane</span></span><br><span data-ttu-id="5aa2c-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-630">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-630">- ActiveView</span></span><br><span data-ttu-id="5aa2c-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-631">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-632">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-633">
         - File</span></span><br><span data-ttu-id="5aa2c-634">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-634">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-635">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-635">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-636">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-636">
         - Selection</span></span><br><span data-ttu-id="5aa2c-637">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-637">
         - Settings</span></span><br><span data-ttu-id="5aa2c-638">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-638">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-639">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="5aa2c-639">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="5aa2c-640">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-640">- Content</span></span><br><span data-ttu-id="5aa2c-641">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-641">
         - TaskPane</span></span><br><span data-ttu-id="5aa2c-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-644">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5aa2c-644">- ActiveView</span></span><br><span data-ttu-id="5aa2c-645">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-645">
         - CompressedFile</span></span><br><span data-ttu-id="5aa2c-646">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-646">
         - DocumentEvents</span></span><br><span data-ttu-id="5aa2c-647">
         - File</span><span class="sxs-lookup"><span data-stu-id="5aa2c-647">
         - File</span></span><br><span data-ttu-id="5aa2c-648">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-648">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5aa2c-649">
         - PdfFile</span></span><br><span data-ttu-id="5aa2c-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-650">
         - Selection</span></span><br><span data-ttu-id="5aa2c-651">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-651">
         - Settings</span></span><br><span data-ttu-id="5aa2c-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-652">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="5aa2c-653">OneNote</span><span class="sxs-lookup"><span data-stu-id="5aa2c-653">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5aa2c-654">Платформа</span><span class="sxs-lookup"><span data-stu-id="5aa2c-654">Platform</span></span></th>
    <th><span data-ttu-id="5aa2c-655">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5aa2c-655">Extension points</span></span></th>
    <th><span data-ttu-id="5aa2c-656">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-656">API requirement sets</span></span></th>
    <th><span data-ttu-id="5aa2c-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-658">Office Online</span><span class="sxs-lookup"><span data-stu-id="5aa2c-658">Office Online</span></span></td>
    <td> <span data-ttu-id="5aa2c-659">- Контент</span><span class="sxs-lookup"><span data-stu-id="5aa2c-659">- Content</span></span><br><span data-ttu-id="5aa2c-660">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-660">
         - TaskPane</span></span><br><span data-ttu-id="5aa2c-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5aa2c-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-664">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5aa2c-664">- DocumentEvents</span></span><br><span data-ttu-id="5aa2c-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="5aa2c-666">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-666">
         - ImageCoercion</span></span><br><span data-ttu-id="5aa2c-667">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="5aa2c-667">
         - Settings</span></span><br><span data-ttu-id="5aa2c-668">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-668">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="5aa2c-669">Project</span><span class="sxs-lookup"><span data-stu-id="5aa2c-669">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5aa2c-670">Платформа</span><span class="sxs-lookup"><span data-stu-id="5aa2c-670">Platform</span></span></th>
    <th><span data-ttu-id="5aa2c-671">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5aa2c-671">Extension points</span></span></th>
    <th><span data-ttu-id="5aa2c-672">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-672">API requirement sets</span></span></th>
    <th><span data-ttu-id="5aa2c-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-674">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-674">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-675">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-675">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5aa2c-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-677">- Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-677">- Selection</span></span><br><span data-ttu-id="5aa2c-678">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-678">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-679">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-679">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-680">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-680">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5aa2c-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-682">- Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-682">- Selection</span></span><br><span data-ttu-id="5aa2c-683">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-683">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5aa2c-684">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="5aa2c-684">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5aa2c-685">- Область задач</span><span class="sxs-lookup"><span data-stu-id="5aa2c-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5aa2c-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5aa2c-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5aa2c-687">- Selection</span><span class="sxs-lookup"><span data-stu-id="5aa2c-687">- Selection</span></span><br><span data-ttu-id="5aa2c-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5aa2c-688">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5aa2c-689">См. также</span><span class="sxs-lookup"><span data-stu-id="5aa2c-689">See also</span></span>

- [<span data-ttu-id="5aa2c-690">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5aa2c-690">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5aa2c-691">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="5aa2c-691">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="5aa2c-692">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="5aa2c-692">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="5aa2c-693">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="5aa2c-693">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
