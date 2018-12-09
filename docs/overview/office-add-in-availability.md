---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 11/07/2018
ms.openlocfilehash: c601eac5ed3fcad76b63fff5ae6eeadb7662c8b7
ms.sourcegitcommit: 0adc31ceaba92cb15dc6430c00fe7a96c107c9de
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2018
ms.locfileid: "27210107"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="bc566-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bc566-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="bc566-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="bc566-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="bc566-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="bc566-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="bc566-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.</span><span class="sxs-lookup"><span data-stu-id="bc566-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="bc566-108">Excel</span><span class="sxs-lookup"><span data-stu-id="bc566-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="bc566-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="bc566-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="bc566-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bc566-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="bc566-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="bc566-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bc566-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="bc566-113">Office Online</span></span></td>
    <td> <span data-ttu-id="bc566-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-114">- TaskPane</span></span><br><span data-ttu-id="bc566-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-115">
        - Content</span></span><br><span data-ttu-id="bc566-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="bc566-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="bc566-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc566-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc566-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc566-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc566-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc566-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc566-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc566-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc566-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc566-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-126">
        - BindingEvents</span></span><br><span data-ttu-id="bc566-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-127">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-128">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-129">
        - File</span></span><br><span data-ttu-id="bc566-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-130">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-132">
        - Selection</span></span><br><span data-ttu-id="bc566-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-133">
        - Settings</span></span><br><span data-ttu-id="bc566-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-134">
        - TableBindings</span></span><br><span data-ttu-id="bc566-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-135">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-136">
        - TextBindings</span></span><br><span data-ttu-id="bc566-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="bc566-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-139">
        - TaskPane</span></span><br><span data-ttu-id="bc566-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="bc566-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-142">
        - BindingEvents</span></span><br><span data-ttu-id="bc566-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-143">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-144">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-145">
        - File</span></span><br><span data-ttu-id="bc566-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-146">
        - ImageCoercion</span></span><br><span data-ttu-id="bc566-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-147">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-149">
        - Selection</span></span><br><span data-ttu-id="bc566-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-150">
        - Settings</span></span><br><span data-ttu-id="bc566-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-151">
        - TableBindings</span></span><br><span data-ttu-id="bc566-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-152">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-153">
        - TextBindings</span></span><br><span data-ttu-id="bc566-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="bc566-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-156">- TaskPane</span></span><br><span data-ttu-id="bc566-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-157">
        - Content</span></span><br><span data-ttu-id="bc566-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc566-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc566-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc566-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc566-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc566-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc566-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc566-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc566-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc566-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc566-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-168">- BindingEvents</span></span><br><span data-ttu-id="bc566-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-169">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-170">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-171">
        - File</span></span><br><span data-ttu-id="bc566-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-172">
        - ImageCoercion</span></span><br><span data-ttu-id="bc566-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-173">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-175">
        - Selection</span></span><br><span data-ttu-id="bc566-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-176">
        - Settings</span></span><br><span data-ttu-id="bc566-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-177">
        - TableBindings</span></span><br><span data-ttu-id="bc566-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-178">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-179">
        - TextBindings</span></span><br><span data-ttu-id="bc566-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-181">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="bc566-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-182">- TaskPane</span></span><br><span data-ttu-id="bc566-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-183">
        - Content</span></span><br><span data-ttu-id="bc566-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc566-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc566-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc566-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc566-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc566-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc566-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc566-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc566-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc566-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc566-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-194">- BindingEvents</span></span><br><span data-ttu-id="bc566-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-195">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-196">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-197">
        - File</span></span><br><span data-ttu-id="bc566-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-198">
        - ImageCoercion</span></span><br><span data-ttu-id="bc566-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-199">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-201">
        - Selection</span></span><br><span data-ttu-id="bc566-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-202">
        - Settings</span></span><br><span data-ttu-id="bc566-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-203">
        - TableBindings</span></span><br><span data-ttu-id="bc566-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-204">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-205">
        - TextBindings</span></span><br><span data-ttu-id="bc566-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-207">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="bc566-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="bc566-208">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-208">- TaskPane</span></span><br><span data-ttu-id="bc566-209">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-209">
        - Content</span></span></td>
    <td><span data-ttu-id="bc566-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc566-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc566-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc566-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc566-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc566-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc566-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc566-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc566-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc566-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-219">- BindingEvents</span></span><br><span data-ttu-id="bc566-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-220">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-221">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-222">
        - File</span></span><br><span data-ttu-id="bc566-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-223">
        - ImageCoercion</span></span><br><span data-ttu-id="bc566-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-224">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-226">
        - Selection</span></span><br><span data-ttu-id="bc566-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-227">
        - Settings</span></span><br><span data-ttu-id="bc566-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-228">
        - TableBindings</span></span><br><span data-ttu-id="bc566-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-229">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-230">
        - TextBindings</span></span><br><span data-ttu-id="bc566-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-232">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="bc566-233">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-233">- TaskPane</span></span><br><span data-ttu-id="bc566-234">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-234">
        - Content</span></span><br><span data-ttu-id="bc566-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc566-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc566-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc566-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc566-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc566-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc566-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc566-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc566-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc566-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc566-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-245">- BindingEvents</span></span><br><span data-ttu-id="bc566-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-246">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-247">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-248">
        - File</span></span><br><span data-ttu-id="bc566-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-249">
        - ImageCoercion</span></span><br><span data-ttu-id="bc566-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-250">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-252">
        - PdfFile</span></span><br><span data-ttu-id="bc566-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-253">
        - Selection</span></span><br><span data-ttu-id="bc566-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-254">
        - Settings</span></span><br><span data-ttu-id="bc566-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-255">
        - TableBindings</span></span><br><span data-ttu-id="bc566-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-256">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-257">
        - TextBindings</span></span><br><span data-ttu-id="bc566-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-259">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="bc566-260">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-260">- TaskPane</span></span><br><span data-ttu-id="bc566-261">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-261">
        - Content</span></span><br><span data-ttu-id="bc566-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc566-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc566-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc566-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc566-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc566-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc566-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc566-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc566-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc566-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc566-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bc566-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-272">- BindingEvents</span></span><br><span data-ttu-id="bc566-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-273">
        - CompressedFile</span></span><br><span data-ttu-id="bc566-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-274">
        - DocumentEvents</span></span><br><span data-ttu-id="bc566-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="bc566-275">
        - File</span></span><br><span data-ttu-id="bc566-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-276">
        - ImageCoercion</span></span><br><span data-ttu-id="bc566-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-277">
        - MatrixBindings</span></span><br><span data-ttu-id="bc566-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc566-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-279">
        - PdfFile</span></span><br><span data-ttu-id="bc566-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-280">
        - Selection</span></span><br><span data-ttu-id="bc566-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-281">
        - Settings</span></span><br><span data-ttu-id="bc566-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-282">
        - TableBindings</span></span><br><span data-ttu-id="bc566-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-283">
        - TableCoercion</span></span><br><span data-ttu-id="bc566-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-284">
        - TextBindings</span></span><br><span data-ttu-id="bc566-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="bc566-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="bc566-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc566-287">Платформа</span><span class="sxs-lookup"><span data-stu-id="bc566-287">Platform</span></span></th>
    <th><span data-ttu-id="bc566-288">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bc566-288">Extension points</span></span></th>
    <th><span data-ttu-id="bc566-289">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc566-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bc566-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="bc566-291">Office Online</span></span></td>
    <td> <span data-ttu-id="bc566-292">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-292">- Mail Read</span></span><br><span data-ttu-id="bc566-293">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bc566-293">
      - Mail Compose</span></span><br><span data-ttu-id="bc566-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc566-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc566-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc566-302">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-303">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-304">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-304">- Mail Read</span></span><br><span data-ttu-id="bc566-305">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bc566-305">
      - Mail Compose</span></span><br><span data-ttu-id="bc566-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="bc566-311">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-312">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-313">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-313">- Mail Read</span></span><br><span data-ttu-id="bc566-314">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bc566-314">
      - Mail Compose</span></span><br><span data-ttu-id="bc566-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bc566-316">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="bc566-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bc566-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc566-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc566-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc566-324">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-325">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-326">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-326">- Mail Read</span></span><br><span data-ttu-id="bc566-327">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bc566-327">
      - Mail Compose</span></span><br><span data-ttu-id="bc566-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bc566-329">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="bc566-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bc566-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc566-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc566-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc566-337">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-338">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="bc566-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="bc566-339">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-339">- Mail Read</span></span><br><span data-ttu-id="bc566-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bc566-346">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-347">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bc566-348">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-348">- Mail Read</span></span><br><span data-ttu-id="bc566-349">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bc566-349">
      - Mail Compose</span></span><br><span data-ttu-id="bc566-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc566-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bc566-357">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-358">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="bc566-359">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-359">- Mail Read</span></span><br><span data-ttu-id="bc566-360">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="bc566-360">
      - Mail Compose</span></span><br><span data-ttu-id="bc566-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc566-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc566-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc566-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc566-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc566-369">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-370">Office для Android</span><span class="sxs-lookup"><span data-stu-id="bc566-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="bc566-371">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="bc566-371">- Mail Read</span></span><br><span data-ttu-id="bc566-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc566-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc566-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc566-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc566-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc566-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc566-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bc566-378">Недоступно</span><span class="sxs-lookup"><span data-stu-id="bc566-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="bc566-379">Word</span><span class="sxs-lookup"><span data-stu-id="bc566-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc566-380">Платформа</span><span class="sxs-lookup"><span data-stu-id="bc566-380">Platform</span></span></th>
    <th><span data-ttu-id="bc566-381">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bc566-381">Extension points</span></span></th>
    <th><span data-ttu-id="bc566-382">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc566-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bc566-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="bc566-384">Office Online</span></span></td>
    <td> <span data-ttu-id="bc566-385">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-385">- TaskPane</span></span><br><span data-ttu-id="bc566-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bc566-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bc566-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bc566-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-391">- BindingEvents</span></span><br><span data-ttu-id="bc566-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-393">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-394">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-394">
         - File</span></span><br><span data-ttu-id="bc566-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-396">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-397">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-400">
         - PdfFile</span></span><br><span data-ttu-id="bc566-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-401">
         - Selection</span></span><br><span data-ttu-id="bc566-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-402">
         - Settings</span></span><br><span data-ttu-id="bc566-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-403">
         - TableBindings</span></span><br><span data-ttu-id="bc566-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-404">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-405">
         - TextBindings</span></span><br><span data-ttu-id="bc566-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-406">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-408">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-409">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc566-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-411">- BindingEvents</span></span><br><span data-ttu-id="bc566-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-412">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-414">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-415">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-415">
         - File</span></span><br><span data-ttu-id="bc566-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-417">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-418">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-421">
         - PdfFile</span></span><br><span data-ttu-id="bc566-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-422">
         - Selection</span></span><br><span data-ttu-id="bc566-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-423">
         - Settings</span></span><br><span data-ttu-id="bc566-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-424">
         - TableBindings</span></span><br><span data-ttu-id="bc566-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-425">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-426">
         - TextBindings</span></span><br><span data-ttu-id="bc566-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-427">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-429">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-430">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-430">- TaskPane</span></span><br><span data-ttu-id="bc566-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bc566-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bc566-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bc566-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-436">- BindingEvents</span></span><br><span data-ttu-id="bc566-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-437">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-439">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-440">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-440">
         - File</span></span><br><span data-ttu-id="bc566-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-442">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-443">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-446">
         - PdfFile</span></span><br><span data-ttu-id="bc566-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-447">
         - Selection</span></span><br><span data-ttu-id="bc566-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-448">
         - Settings</span></span><br><span data-ttu-id="bc566-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-449">
         - TableBindings</span></span><br><span data-ttu-id="bc566-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-450">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-451">
         - TextBindings</span></span><br><span data-ttu-id="bc566-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-452">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-454">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-455">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-455">- TaskPane</span></span><br><span data-ttu-id="bc566-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bc566-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bc566-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bc566-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-461">- BindingEvents</span></span><br><span data-ttu-id="bc566-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-462">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-464">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-465">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-465">
         - File</span></span><br><span data-ttu-id="bc566-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-467">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-468">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-471">
         - PdfFile</span></span><br><span data-ttu-id="bc566-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-472">
         - Selection</span></span><br><span data-ttu-id="bc566-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-473">
         - Settings</span></span><br><span data-ttu-id="bc566-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-474">
         - TableBindings</span></span><br><span data-ttu-id="bc566-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-475">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-476">
         - TextBindings</span></span><br><span data-ttu-id="bc566-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-477">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-479">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="bc566-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="bc566-480">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc566-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bc566-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bc566-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bc566-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bc566-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bc566-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-485">- BindingEvents</span></span><br><span data-ttu-id="bc566-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-486">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-488">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-489">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-489">
         - File</span></span><br><span data-ttu-id="bc566-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-491">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-492">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-495">
         - PdfFile</span></span><br><span data-ttu-id="bc566-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-496">
         - Selection</span></span><br><span data-ttu-id="bc566-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-497">
         - Settings</span></span><br><span data-ttu-id="bc566-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-498">
         - TableBindings</span></span><br><span data-ttu-id="bc566-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-499">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-500">
         - TextBindings</span></span><br><span data-ttu-id="bc566-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-501">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-503">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bc566-504">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-504">- TaskPane</span></span><br><span data-ttu-id="bc566-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bc566-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bc566-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bc566-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bc566-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bc566-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-510">- BindingEvents</span></span><br><span data-ttu-id="bc566-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-511">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-513">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-514">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-514">
         - File</span></span><br><span data-ttu-id="bc566-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-516">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-517">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-520">
         - PdfFile</span></span><br><span data-ttu-id="bc566-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-521">
         - Selection</span></span><br><span data-ttu-id="bc566-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-522">
         - Settings</span></span><br><span data-ttu-id="bc566-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-523">
         - TableBindings</span></span><br><span data-ttu-id="bc566-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-524">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-525">
         - TextBindings</span></span><br><span data-ttu-id="bc566-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-526">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-528">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="bc566-529">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-529">- TaskPane</span></span><br><span data-ttu-id="bc566-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bc566-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc566-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bc566-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc566-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bc566-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bc566-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bc566-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-535">- BindingEvents</span></span><br><span data-ttu-id="bc566-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-536">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc566-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc566-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-538">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="bc566-539">
         - File</span></span><br><span data-ttu-id="bc566-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-541">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-542">
         - MatrixBindings</span></span><br><span data-ttu-id="bc566-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc566-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc566-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-545">
         - PdfFile</span></span><br><span data-ttu-id="bc566-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-546">
         - Selection</span></span><br><span data-ttu-id="bc566-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bc566-547">
         - Settings</span></span><br><span data-ttu-id="bc566-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-548">
         - TableBindings</span></span><br><span data-ttu-id="bc566-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-549">
         - TableCoercion</span></span><br><span data-ttu-id="bc566-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc566-550">
         - TextBindings</span></span><br><span data-ttu-id="bc566-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-551">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc566-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="bc566-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bc566-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc566-554">Платформа</span><span class="sxs-lookup"><span data-stu-id="bc566-554">Platform</span></span></th>
    <th><span data-ttu-id="bc566-555">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bc566-555">Extension points</span></span></th>
    <th><span data-ttu-id="bc566-556">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc566-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bc566-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="bc566-558">Office Online</span></span></td>
    <td> <span data-ttu-id="bc566-559">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-559">- Content</span></span><br><span data-ttu-id="bc566-560">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-560">
         - TaskPane</span></span><br><span data-ttu-id="bc566-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-563">- ActiveView</span></span><br><span data-ttu-id="bc566-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-564">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-565">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-566">
         - File</span></span><br><span data-ttu-id="bc566-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-567">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-568">
         - PdfFile</span></span><br><span data-ttu-id="bc566-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-569">
         - Selection</span></span><br><span data-ttu-id="bc566-570">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-570">
         - Settings</span></span><br><span data-ttu-id="bc566-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-572">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-573">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-573">- Content</span></span><br><span data-ttu-id="bc566-574">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="bc566-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bc566-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bc566-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-576">- ActiveView</span></span><br><span data-ttu-id="bc566-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-577">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-578">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-579">
         - File</span></span><br><span data-ttu-id="bc566-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-580">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-581">
         - PdfFile</span></span><br><span data-ttu-id="bc566-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-582">
         - Selection</span></span><br><span data-ttu-id="bc566-583">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-583">
         - Settings</span></span><br><span data-ttu-id="bc566-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-586">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-586">- Content</span></span><br><span data-ttu-id="bc566-587">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-587">
         - TaskPane</span></span><br><span data-ttu-id="bc566-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-590">- ActiveView</span></span><br><span data-ttu-id="bc566-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-591">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-592">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-593">
         - File</span></span><br><span data-ttu-id="bc566-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-594">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-595">
         - PdfFile</span></span><br><span data-ttu-id="bc566-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-596">
         - Selection</span></span><br><span data-ttu-id="bc566-597">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-597">
         - Settings</span></span><br><span data-ttu-id="bc566-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-599">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-600">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-600">- Content</span></span><br><span data-ttu-id="bc566-601">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-601">
         - TaskPane</span></span><br><span data-ttu-id="bc566-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-604">- ActiveView</span></span><br><span data-ttu-id="bc566-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-605">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-606">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-607">
         - File</span></span><br><span data-ttu-id="bc566-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-608">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-609">
         - PdfFile</span></span><br><span data-ttu-id="bc566-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-610">
         - Selection</span></span><br><span data-ttu-id="bc566-611">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-611">
         - Settings</span></span><br><span data-ttu-id="bc566-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-613">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="bc566-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="bc566-614">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-614">- Content</span></span><br><span data-ttu-id="bc566-615">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="bc566-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="bc566-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-617">- ActiveView</span></span><br><span data-ttu-id="bc566-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-618">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-619">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-620">
         - File</span></span><br><span data-ttu-id="bc566-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-621">
         - PdfFile</span></span><br><span data-ttu-id="bc566-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-622">
         - Selection</span></span><br><span data-ttu-id="bc566-623">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-623">
         - Settings</span></span><br><span data-ttu-id="bc566-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-624">
         - TextCoercion</span></span><br><span data-ttu-id="bc566-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-626">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bc566-627">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-627">- Content</span></span><br><span data-ttu-id="bc566-628">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-628">
         - TaskPane</span></span><br><span data-ttu-id="bc566-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-631">- ActiveView</span></span><br><span data-ttu-id="bc566-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-632">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-633">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-634">
         - File</span></span><br><span data-ttu-id="bc566-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-635">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-636">
         - PdfFile</span></span><br><span data-ttu-id="bc566-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-637">
         - Selection</span></span><br><span data-ttu-id="bc566-638">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-638">
         - Settings</span></span><br><span data-ttu-id="bc566-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-640">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="bc566-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="bc566-641">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-641">- Content</span></span><br><span data-ttu-id="bc566-642">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-642">
         - TaskPane</span></span><br><span data-ttu-id="bc566-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc566-645">- ActiveView</span></span><br><span data-ttu-id="bc566-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc566-646">
         - CompressedFile</span></span><br><span data-ttu-id="bc566-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-647">
         - DocumentEvents</span></span><br><span data-ttu-id="bc566-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="bc566-648">
         - File</span></span><br><span data-ttu-id="bc566-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-649">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc566-650">
         - PdfFile</span></span><br><span data-ttu-id="bc566-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-651">
         - Selection</span></span><br><span data-ttu-id="bc566-652">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-652">
         - Settings</span></span><br><span data-ttu-id="bc566-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="bc566-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="bc566-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc566-655">Платформа</span><span class="sxs-lookup"><span data-stu-id="bc566-655">Platform</span></span></th>
    <th><span data-ttu-id="bc566-656">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bc566-656">Extension points</span></span></th>
    <th><span data-ttu-id="bc566-657">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc566-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></span><span class="sxs-lookup"><span data-stu-id="bc566-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="bc566-659">Office Online</span></span></td>
    <td> <span data-ttu-id="bc566-660">- Контент</span><span class="sxs-lookup"><span data-stu-id="bc566-660">- Content</span></span><br><span data-ttu-id="bc566-661">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-661">
         - TaskPane</span></span><br><span data-ttu-id="bc566-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="bc566-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc566-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="bc566-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc566-665">- DocumentEvents</span></span><br><span data-ttu-id="bc566-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc566-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-667">
         - ImageCoercion</span></span><br><span data-ttu-id="bc566-668">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="bc566-668">
         - Settings</span></span><br><span data-ttu-id="bc566-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="bc566-670">Project</span><span class="sxs-lookup"><span data-stu-id="bc566-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc566-671">Платформа</span><span class="sxs-lookup"><span data-stu-id="bc566-671">Platform</span></span></th>
    <th><span data-ttu-id="bc566-672">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="bc566-672">Extension points</span></span></th>
    <th><span data-ttu-id="bc566-673">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc566-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Стандартные API</b></a></span><span class="sxs-lookup"><span data-stu-id="bc566-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-675">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-676">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc566-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-678">- Selection</span></span><br><span data-ttu-id="bc566-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-680">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-681">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc566-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-683">- Selection</span></span><br><span data-ttu-id="bc566-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc566-685">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="bc566-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bc566-686">- Область задач</span><span class="sxs-lookup"><span data-stu-id="bc566-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc566-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc566-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc566-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="bc566-688">- Selection</span></span><br><span data-ttu-id="bc566-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc566-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="bc566-690">См. также</span><span class="sxs-lookup"><span data-stu-id="bc566-690">See also</span></span>

- [<span data-ttu-id="bc566-691">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bc566-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="bc566-692">Стандартные наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="bc566-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="bc566-693">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="bc566-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="bc566-694">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="bc566-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
