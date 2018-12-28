---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 11/07/2018
ms.openlocfilehash: 9490fca9663737e2397de159169b545e3900289f
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458043"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4c21a-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="4c21a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4c21a-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="4c21a-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="4c21a-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="4c21a-105">The following tables contain the available platforms, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4c21a-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="4c21a-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="4c21a-108">Excel</span><span class="sxs-lookup"><span data-stu-id="4c21a-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4c21a-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="4c21a-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4c21a-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4c21a-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4c21a-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4c21a-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4c21a-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4c21a-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c21a-113">Office Online</span></span></td>
    <td> <span data-ttu-id="4c21a-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-114">- TaskPane</span></span><br><span data-ttu-id="4c21a-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-115">
        - Content</span></span><br><span data-ttu-id="4c21a-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="4c21a-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4c21a-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c21a-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c21a-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c21a-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c21a-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c21a-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c21a-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4c21a-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4c21a-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-126">
        - BindingEvents</span></span><br><span data-ttu-id="4c21a-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-127">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-128">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-129">
        - File</span></span><br><span data-ttu-id="4c21a-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-130">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-132">
        - Selection</span></span><br><span data-ttu-id="4c21a-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-133">
        - Settings</span></span><br><span data-ttu-id="4c21a-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-134">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-135">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-136">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-138">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4c21a-139">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-139">
        - TaskPane</span></span><br><span data-ttu-id="4c21a-140">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4c21a-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-142">
        - BindingEvents</span></span><br><span data-ttu-id="4c21a-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-143">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-144">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-145">
        - File</span></span><br><span data-ttu-id="4c21a-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-146">
        - ImageCoercion</span></span><br><span data-ttu-id="4c21a-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-147">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-149">
        - Selection</span></span><br><span data-ttu-id="4c21a-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-150">
        - Settings</span></span><br><span data-ttu-id="4c21a-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-151">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-152">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-153">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-155">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4c21a-156">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-156">- TaskPane</span></span><br><span data-ttu-id="4c21a-157">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-157">
        - Content</span></span><br><span data-ttu-id="4c21a-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4c21a-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c21a-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c21a-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c21a-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c21a-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c21a-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c21a-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4c21a-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4c21a-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-168">- BindingEvents</span></span><br><span data-ttu-id="4c21a-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-169">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-170">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-171">
        - File</span></span><br><span data-ttu-id="4c21a-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-172">
        - ImageCoercion</span></span><br><span data-ttu-id="4c21a-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-173">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-175">
        - Selection</span></span><br><span data-ttu-id="4c21a-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-176">
        - Settings</span></span><br><span data-ttu-id="4c21a-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-177">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-178">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-179">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-181">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="4c21a-182">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-182">- TaskPane</span></span><br><span data-ttu-id="4c21a-183">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-183">
        - Content</span></span><br><span data-ttu-id="4c21a-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4c21a-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c21a-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c21a-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c21a-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c21a-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c21a-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c21a-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4c21a-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4c21a-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-194">- BindingEvents</span></span><br><span data-ttu-id="4c21a-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-195">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-196">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-197">
        - File</span></span><br><span data-ttu-id="4c21a-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-198">
        - ImageCoercion</span></span><br><span data-ttu-id="4c21a-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-199">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-201">
        - Selection</span></span><br><span data-ttu-id="4c21a-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-202">
        - Settings</span></span><br><span data-ttu-id="4c21a-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-203">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-204">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-205">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-207">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="4c21a-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="4c21a-208">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-208">- TaskPane</span></span><br><span data-ttu-id="4c21a-209">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-209">
        - Content</span></span></td>
    <td><span data-ttu-id="4c21a-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c21a-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c21a-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c21a-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c21a-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c21a-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c21a-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4c21a-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4c21a-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-219">- BindingEvents</span></span><br><span data-ttu-id="4c21a-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-220">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-221">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-222">
        - File</span></span><br><span data-ttu-id="4c21a-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-223">
        - ImageCoercion</span></span><br><span data-ttu-id="4c21a-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-224">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-226">
        - Selection</span></span><br><span data-ttu-id="4c21a-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-227">
        - Settings</span></span><br><span data-ttu-id="4c21a-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-228">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-229">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-230">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-232">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4c21a-233">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-233">- TaskPane</span></span><br><span data-ttu-id="4c21a-234">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-234">
        - Content</span></span><br><span data-ttu-id="4c21a-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4c21a-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c21a-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c21a-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c21a-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c21a-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c21a-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c21a-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4c21a-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4c21a-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-245">- BindingEvents</span></span><br><span data-ttu-id="4c21a-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-246">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-247">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-248">
        - File</span></span><br><span data-ttu-id="4c21a-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-249">
        - ImageCoercion</span></span><br><span data-ttu-id="4c21a-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-250">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-252">
        - PdfFile</span></span><br><span data-ttu-id="4c21a-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-253">
        - Selection</span></span><br><span data-ttu-id="4c21a-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-254">
        - Settings</span></span><br><span data-ttu-id="4c21a-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-255">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-256">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-257">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-259">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="4c21a-260">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-260">- TaskPane</span></span><br><span data-ttu-id="4c21a-261">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-261">
        - Content</span></span><br><span data-ttu-id="4c21a-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4c21a-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4c21a-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4c21a-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4c21a-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4c21a-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4c21a-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4c21a-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4c21a-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4c21a-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4c21a-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-272">- BindingEvents</span></span><br><span data-ttu-id="4c21a-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-273">
        - CompressedFile</span></span><br><span data-ttu-id="4c21a-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-274">
        - DocumentEvents</span></span><br><span data-ttu-id="4c21a-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-275">
        - File</span></span><br><span data-ttu-id="4c21a-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-276">
        - ImageCoercion</span></span><br><span data-ttu-id="4c21a-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-277">
        - MatrixBindings</span></span><br><span data-ttu-id="4c21a-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-279">
        - PdfFile</span></span><br><span data-ttu-id="4c21a-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-280">
        - Selection</span></span><br><span data-ttu-id="4c21a-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-281">
        - Settings</span></span><br><span data-ttu-id="4c21a-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-282">
        - TableBindings</span></span><br><span data-ttu-id="4c21a-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-283">
        - TableCoercion</span></span><br><span data-ttu-id="4c21a-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-284">
        - TextBindings</span></span><br><span data-ttu-id="4c21a-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="4c21a-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="4c21a-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c21a-287">Платформа</span><span class="sxs-lookup"><span data-stu-id="4c21a-287">Platform</span></span></th>
    <th><span data-ttu-id="4c21a-288">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4c21a-288">Extension points</span></span></th>
    <th><span data-ttu-id="4c21a-289">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4c21a-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c21a-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4c21a-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c21a-291">Office Online</span></span></td>
    <td> <span data-ttu-id="4c21a-292">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-292">- Mail Read</span></span><br><span data-ttu-id="4c21a-293">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-293">
      - Mail Compose</span></span><br><span data-ttu-id="4c21a-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c21a-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4c21a-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4c21a-302">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-303">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-304">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-304">- Mail Read</span></span><br><span data-ttu-id="4c21a-305">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-305">
      - Mail Compose</span></span><br><span data-ttu-id="4c21a-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4c21a-311">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-312">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-313">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-313">- Mail Read</span></span><br><span data-ttu-id="4c21a-314">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-314">
      - Mail Compose</span></span><br><span data-ttu-id="4c21a-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4c21a-316">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="4c21a-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4c21a-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c21a-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4c21a-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4c21a-324">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-325">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-326">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-326">- Mail Read</span></span><br><span data-ttu-id="4c21a-327">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-327">
      - Mail Compose</span></span><br><span data-ttu-id="4c21a-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4c21a-329">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="4c21a-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4c21a-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c21a-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4c21a-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4c21a-337">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-338">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="4c21a-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4c21a-339">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-339">- Mail Read</span></span><br><span data-ttu-id="4c21a-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4c21a-346">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-347">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4c21a-348">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-348">- Mail Read</span></span><br><span data-ttu-id="4c21a-349">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-349">
      - Mail Compose</span></span><br><span data-ttu-id="4c21a-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c21a-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4c21a-357">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-358">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4c21a-359">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-359">- Mail Read</span></span><br><span data-ttu-id="4c21a-360">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-360">
      - Mail Compose</span></span><br><span data-ttu-id="4c21a-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4c21a-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4c21a-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4c21a-369">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-370">Office для Android</span><span class="sxs-lookup"><span data-stu-id="4c21a-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="4c21a-371">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4c21a-371">- Mail Read</span></span><br><span data-ttu-id="4c21a-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4c21a-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4c21a-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4c21a-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4c21a-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4c21a-378">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4c21a-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="4c21a-379">Word</span><span class="sxs-lookup"><span data-stu-id="4c21a-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c21a-380">Платформа</span><span class="sxs-lookup"><span data-stu-id="4c21a-380">Platform</span></span></th>
    <th><span data-ttu-id="4c21a-381">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4c21a-381">Extension points</span></span></th>
    <th><span data-ttu-id="4c21a-382">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4c21a-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c21a-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4c21a-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c21a-384">Office Online</span></span></td>
    <td> <span data-ttu-id="4c21a-385">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-385">- TaskPane</span></span><br><span data-ttu-id="4c21a-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c21a-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c21a-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c21a-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-391">- BindingEvents</span></span><br><span data-ttu-id="4c21a-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-393">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-394">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-394">
         - File</span></span><br><span data-ttu-id="4c21a-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-396">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-397">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-400">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-401">
         - Selection</span></span><br><span data-ttu-id="4c21a-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-402">
         - Settings</span></span><br><span data-ttu-id="4c21a-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-403">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-404">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-405">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-406">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-408">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-409">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4c21a-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-411">- BindingEvents</span></span><br><span data-ttu-id="4c21a-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-412">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-414">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-415">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-415">
         - File</span></span><br><span data-ttu-id="4c21a-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-417">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-418">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-421">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-422">
         - Selection</span></span><br><span data-ttu-id="4c21a-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-423">
         - Settings</span></span><br><span data-ttu-id="4c21a-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-424">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-425">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-426">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-427">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-429">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-430">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-430">- TaskPane</span></span><br><span data-ttu-id="4c21a-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c21a-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c21a-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c21a-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-436">- BindingEvents</span></span><br><span data-ttu-id="4c21a-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-437">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-439">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-440">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-440">
         - File</span></span><br><span data-ttu-id="4c21a-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-442">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-443">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-446">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-447">
         - Selection</span></span><br><span data-ttu-id="4c21a-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-448">
         - Settings</span></span><br><span data-ttu-id="4c21a-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-449">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-450">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-451">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-452">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-454">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-455">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-455">- TaskPane</span></span><br><span data-ttu-id="4c21a-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c21a-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c21a-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c21a-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-461">- BindingEvents</span></span><br><span data-ttu-id="4c21a-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-462">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-464">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-465">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-465">
         - File</span></span><br><span data-ttu-id="4c21a-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-467">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-468">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-471">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-472">
         - Selection</span></span><br><span data-ttu-id="4c21a-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-473">
         - Settings</span></span><br><span data-ttu-id="4c21a-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-474">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-475">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-476">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-477">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-479">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="4c21a-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="4c21a-480">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4c21a-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c21a-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c21a-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c21a-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c21a-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c21a-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-485">- BindingEvents</span></span><br><span data-ttu-id="4c21a-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-486">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-488">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-489">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-489">
         - File</span></span><br><span data-ttu-id="4c21a-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-491">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-492">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-495">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-496">
         - Selection</span></span><br><span data-ttu-id="4c21a-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-497">
         - Settings</span></span><br><span data-ttu-id="4c21a-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-498">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-499">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-500">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-501">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-503">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4c21a-504">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-504">- TaskPane</span></span><br><span data-ttu-id="4c21a-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c21a-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c21a-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c21a-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c21a-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c21a-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-510">- BindingEvents</span></span><br><span data-ttu-id="4c21a-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-511">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-513">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-514">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-514">
         - File</span></span><br><span data-ttu-id="4c21a-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-516">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-517">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-520">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-521">
         - Selection</span></span><br><span data-ttu-id="4c21a-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-522">
         - Settings</span></span><br><span data-ttu-id="4c21a-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-523">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-524">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-525">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-526">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-528">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4c21a-529">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-529">- TaskPane</span></span><br><span data-ttu-id="4c21a-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4c21a-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4c21a-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4c21a-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c21a-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c21a-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-535">- BindingEvents</span></span><br><span data-ttu-id="4c21a-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-536">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4c21a-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="4c21a-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-538">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4c21a-539">
         - File</span></span><br><span data-ttu-id="4c21a-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-541">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-542">
         - MatrixBindings</span></span><br><span data-ttu-id="4c21a-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="4c21a-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4c21a-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-545">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-546">
         - Selection</span></span><br><span data-ttu-id="4c21a-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4c21a-547">
         - Settings</span></span><br><span data-ttu-id="4c21a-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-548">
         - TableBindings</span></span><br><span data-ttu-id="4c21a-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-549">
         - TableCoercion</span></span><br><span data-ttu-id="4c21a-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4c21a-550">
         - TextBindings</span></span><br><span data-ttu-id="4c21a-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-551">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4c21a-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4c21a-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c21a-554">Платформа</span><span class="sxs-lookup"><span data-stu-id="4c21a-554">Platform</span></span></th>
    <th><span data-ttu-id="4c21a-555">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4c21a-555">Extension points</span></span></th>
    <th><span data-ttu-id="4c21a-556">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4c21a-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c21a-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4c21a-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c21a-558">Office Online</span></span></td>
    <td> <span data-ttu-id="4c21a-559">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-559">- Content</span></span><br><span data-ttu-id="4c21a-560">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-560">
         - TaskPane</span></span><br><span data-ttu-id="4c21a-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-563">- ActiveView</span></span><br><span data-ttu-id="4c21a-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-564">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-565">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-566">
         - File</span></span><br><span data-ttu-id="4c21a-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-567">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-568">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-569">
         - Selection</span></span><br><span data-ttu-id="4c21a-570">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-570">
         - Settings</span></span><br><span data-ttu-id="4c21a-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-572">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-573">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-573">- Content</span></span><br><span data-ttu-id="4c21a-574">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4c21a-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4c21a-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4c21a-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-576">- ActiveView</span></span><br><span data-ttu-id="4c21a-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-577">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-578">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-579">
         - File</span></span><br><span data-ttu-id="4c21a-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-580">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-581">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-582">
         - Selection</span></span><br><span data-ttu-id="4c21a-583">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-583">
         - Settings</span></span><br><span data-ttu-id="4c21a-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-586">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-586">- Content</span></span><br><span data-ttu-id="4c21a-587">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-587">
         - TaskPane</span></span><br><span data-ttu-id="4c21a-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-590">- ActiveView</span></span><br><span data-ttu-id="4c21a-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-591">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-592">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-593">
         - File</span></span><br><span data-ttu-id="4c21a-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-594">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-595">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-596">
         - Selection</span></span><br><span data-ttu-id="4c21a-597">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-597">
         - Settings</span></span><br><span data-ttu-id="4c21a-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-599">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-600">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-600">- Content</span></span><br><span data-ttu-id="4c21a-601">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-601">
         - TaskPane</span></span><br><span data-ttu-id="4c21a-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-604">- ActiveView</span></span><br><span data-ttu-id="4c21a-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-605">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-606">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-607">
         - File</span></span><br><span data-ttu-id="4c21a-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-608">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-609">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-610">
         - Selection</span></span><br><span data-ttu-id="4c21a-611">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-611">
         - Settings</span></span><br><span data-ttu-id="4c21a-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-613">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="4c21a-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="4c21a-614">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-614">- Content</span></span><br><span data-ttu-id="4c21a-615">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4c21a-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4c21a-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-617">- ActiveView</span></span><br><span data-ttu-id="4c21a-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-618">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-619">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-620">
         - File</span></span><br><span data-ttu-id="4c21a-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-621">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-622">
         - Selection</span></span><br><span data-ttu-id="4c21a-623">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-623">
         - Settings</span></span><br><span data-ttu-id="4c21a-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-624">
         - TextCoercion</span></span><br><span data-ttu-id="4c21a-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-626">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4c21a-627">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-627">- Content</span></span><br><span data-ttu-id="4c21a-628">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-628">
         - TaskPane</span></span><br><span data-ttu-id="4c21a-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-631">- ActiveView</span></span><br><span data-ttu-id="4c21a-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-632">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-633">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-634">
         - File</span></span><br><span data-ttu-id="4c21a-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-635">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-636">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-637">
         - Selection</span></span><br><span data-ttu-id="4c21a-638">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-638">
         - Settings</span></span><br><span data-ttu-id="4c21a-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-640">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4c21a-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4c21a-641">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-641">- Content</span></span><br><span data-ttu-id="4c21a-642">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-642">
         - TaskPane</span></span><br><span data-ttu-id="4c21a-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4c21a-645">- ActiveView</span></span><br><span data-ttu-id="4c21a-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-646">
         - CompressedFile</span></span><br><span data-ttu-id="4c21a-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-647">
         - DocumentEvents</span></span><br><span data-ttu-id="4c21a-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="4c21a-648">
         - File</span></span><br><span data-ttu-id="4c21a-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-649">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4c21a-650">
         - PdfFile</span></span><br><span data-ttu-id="4c21a-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-651">
         - Selection</span></span><br><span data-ttu-id="4c21a-652">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-652">
         - Settings</span></span><br><span data-ttu-id="4c21a-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="4c21a-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="4c21a-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c21a-655">Платформа</span><span class="sxs-lookup"><span data-stu-id="4c21a-655">Platform</span></span></th>
    <th><span data-ttu-id="4c21a-656">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4c21a-656">Extension points</span></span></th>
    <th><span data-ttu-id="4c21a-657">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4c21a-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c21a-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4c21a-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="4c21a-659">Office Online</span></span></td>
    <td> <span data-ttu-id="4c21a-660">- Контент</span><span class="sxs-lookup"><span data-stu-id="4c21a-660">- Content</span></span><br><span data-ttu-id="4c21a-661">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-661">
         - TaskPane</span></span><br><span data-ttu-id="4c21a-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4c21a-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4c21a-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4c21a-665">- DocumentEvents</span></span><br><span data-ttu-id="4c21a-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="4c21a-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-667">
         - ImageCoercion</span></span><br><span data-ttu-id="4c21a-668">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4c21a-668">
         - Settings</span></span><br><span data-ttu-id="4c21a-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4c21a-670">Project</span><span class="sxs-lookup"><span data-stu-id="4c21a-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4c21a-671">Платформа</span><span class="sxs-lookup"><span data-stu-id="4c21a-671">Platform</span></span></th>
    <th><span data-ttu-id="4c21a-672">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4c21a-672">Extension points</span></span></th>
    <th><span data-ttu-id="4c21a-673">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4c21a-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="4c21a-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4c21a-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-675">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-676">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4c21a-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-678">- Selection</span></span><br><span data-ttu-id="4c21a-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-680">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-681">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4c21a-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-683">- Selection</span></span><br><span data-ttu-id="4c21a-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4c21a-685">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4c21a-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4c21a-686">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4c21a-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4c21a-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4c21a-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4c21a-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="4c21a-688">- Selection</span></span><br><span data-ttu-id="4c21a-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4c21a-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4c21a-690">См. также</span><span class="sxs-lookup"><span data-stu-id="4c21a-690">See also</span></span>

- [<span data-ttu-id="4c21a-691">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="4c21a-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4c21a-692">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="4c21a-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4c21a-693">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="4c21a-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4c21a-694">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="4c21a-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
