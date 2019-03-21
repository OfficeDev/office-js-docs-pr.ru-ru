---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: fe5b1d1278d2c14192fb6fd212f24bb08571d35d
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691127"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="0b412-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0b412-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="0b412-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="0b412-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="0b412-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="0b412-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="0b412-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="0b412-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="0b412-108">Номер сборки для единовременной покупки Office 2019 — 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="0b412-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="0b412-109">Excel</span><span class="sxs-lookup"><span data-stu-id="0b412-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="0b412-110">Платформа</span><span class="sxs-lookup"><span data-stu-id="0b412-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="0b412-111">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0b412-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="0b412-112">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0b412-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="0b412-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b412-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b412-114">Office Online</span></span></td>
    <td> <span data-ttu-id="0b412-115">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-115">- TaskPane</span></span><br><span data-ttu-id="0b412-116">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-116">
        - Content</span></span><br><span data-ttu-id="0b412-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="0b412-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="0b412-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b412-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b412-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b412-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b412-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b412-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b412-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b412-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b412-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b412-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-127">
        - BindingEvents</span></span><br><span data-ttu-id="0b412-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-128">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-129">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-130">
        - File</span></span><br><span data-ttu-id="0b412-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-131">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-133">
        - Selection</span></span><br><span data-ttu-id="0b412-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-134">
        - Settings</span></span><br><span data-ttu-id="0b412-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-135">
        - TableBindings</span></span><br><span data-ttu-id="0b412-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-136">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-137">
        - TextBindings</span></span><br><span data-ttu-id="0b412-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-139">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-140">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-140">- TaskPane</span></span><br><span data-ttu-id="0b412-141">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-141">
        - Content</span></span><br><span data-ttu-id="0b412-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="0b412-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="0b412-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b412-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b412-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b412-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b412-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b412-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b412-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b412-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b412-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b412-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-152">
        - BindingEvents</span></span><br><span data-ttu-id="0b412-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-153">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-154">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-155">
        - File</span></span><br><span data-ttu-id="0b412-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-156">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-158">
        - Selection</span></span><br><span data-ttu-id="0b412-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-159">
        - Settings</span></span><br><span data-ttu-id="0b412-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-160">
        - TableBindings</span></span><br><span data-ttu-id="0b412-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-161">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-162">
        - TextBindings</span></span><br><span data-ttu-id="0b412-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-164">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="0b412-165">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-165">- TaskPane</span></span><br><span data-ttu-id="0b412-166">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-166">
        - Content</span></span><br><span data-ttu-id="0b412-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0b412-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b412-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b412-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b412-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b412-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b412-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b412-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b412-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b412-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b412-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-177">- BindingEvents</span></span><br><span data-ttu-id="0b412-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-178">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-179">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-180">
        - File</span></span><br><span data-ttu-id="0b412-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-181">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-182">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-184">
        - Selection</span></span><br><span data-ttu-id="0b412-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-185">
        - Settings</span></span><br><span data-ttu-id="0b412-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-186">
        - TableBindings</span></span><br><span data-ttu-id="0b412-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-187">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-188">
        - TextBindings</span></span><br><span data-ttu-id="0b412-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-190">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="0b412-191">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-191">- TaskPane</span></span><br><span data-ttu-id="0b412-192">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-192">
        - Content</span></span></td>
    <td><span data-ttu-id="0b412-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="0b412-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-195">- BindingEvents</span></span><br><span data-ttu-id="0b412-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-196">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-197">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-198">
        - File</span></span><br><span data-ttu-id="0b412-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-199">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-200">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-202">
        - Selection</span></span><br><span data-ttu-id="0b412-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-203">
        - Settings</span></span><br><span data-ttu-id="0b412-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-204">
        - TableBindings</span></span><br><span data-ttu-id="0b412-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-205">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-206">
        - TextBindings</span></span><br><span data-ttu-id="0b412-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-208">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="0b412-209">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-209">
        - TaskPane</span></span><br><span data-ttu-id="0b412-210">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="0b412-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b412-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="0b412-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-212">
        - BindingEvents</span></span><br><span data-ttu-id="0b412-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-213">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-214">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-215">
        - File</span></span><br><span data-ttu-id="0b412-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-216">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-217">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-219">
        - Selection</span></span><br><span data-ttu-id="0b412-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-220">
        - Settings</span></span><br><span data-ttu-id="0b412-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-221">
        - TableBindings</span></span><br><span data-ttu-id="0b412-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-222">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-223">
        - TextBindings</span></span><br><span data-ttu-id="0b412-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-225">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="0b412-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="0b412-226">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-226">- TaskPane</span></span><br><span data-ttu-id="0b412-227">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-227">
        - Content</span></span></td>
    <td><span data-ttu-id="0b412-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b412-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b412-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b412-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b412-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b412-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b412-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b412-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b412-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b412-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-237">- BindingEvents</span></span><br><span data-ttu-id="0b412-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-238">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-239">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-240">
        - File</span></span><br><span data-ttu-id="0b412-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-241">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-242">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-244">
        - Selection</span></span><br><span data-ttu-id="0b412-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-245">
        - Settings</span></span><br><span data-ttu-id="0b412-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-246">
        - TableBindings</span></span><br><span data-ttu-id="0b412-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-247">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-248">
        - TextBindings</span></span><br><span data-ttu-id="0b412-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-250">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="0b412-251">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-251">- TaskPane</span></span><br><span data-ttu-id="0b412-252">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-252">
        - Content</span></span><br><span data-ttu-id="0b412-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0b412-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b412-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b412-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b412-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b412-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b412-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b412-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b412-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b412-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b412-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-263">- BindingEvents</span></span><br><span data-ttu-id="0b412-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-264">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-265">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-266">
        - File</span></span><br><span data-ttu-id="0b412-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-267">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-268">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-270">
        - PdfFile</span></span><br><span data-ttu-id="0b412-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-271">
        - Selection</span></span><br><span data-ttu-id="0b412-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-272">
        - Settings</span></span><br><span data-ttu-id="0b412-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-273">
        - TableBindings</span></span><br><span data-ttu-id="0b412-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-274">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-275">
        - TextBindings</span></span><br><span data-ttu-id="0b412-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-277">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="0b412-278">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-278">- TaskPane</span></span><br><span data-ttu-id="0b412-279">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-279">
        - Content</span></span><br><span data-ttu-id="0b412-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="0b412-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="0b412-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="0b412-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="0b412-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="0b412-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="0b412-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="0b412-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="0b412-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="0b412-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="0b412-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-290">- BindingEvents</span></span><br><span data-ttu-id="0b412-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-291">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-292">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-293">
        - File</span></span><br><span data-ttu-id="0b412-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-294">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-295">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-297">
        - PdfFile</span></span><br><span data-ttu-id="0b412-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-298">
        - Selection</span></span><br><span data-ttu-id="0b412-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-299">
        - Settings</span></span><br><span data-ttu-id="0b412-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-300">
        - TableBindings</span></span><br><span data-ttu-id="0b412-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-301">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-302">
        - TextBindings</span></span><br><span data-ttu-id="0b412-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-304">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="0b412-305">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-305">- TaskPane</span></span><br><span data-ttu-id="0b412-306">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-306">
        - Content</span></span></td>
    <td><span data-ttu-id="0b412-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="0b412-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="0b412-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-309">- BindingEvents</span></span><br><span data-ttu-id="0b412-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-310">
        - CompressedFile</span></span><br><span data-ttu-id="0b412-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-311">
        - DocumentEvents</span></span><br><span data-ttu-id="0b412-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="0b412-312">
        - File</span></span><br><span data-ttu-id="0b412-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-313">
        - ImageCoercion</span></span><br><span data-ttu-id="0b412-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-314">
        - MatrixBindings</span></span><br><span data-ttu-id="0b412-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="0b412-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-316">
        - PdfFile</span></span><br><span data-ttu-id="0b412-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-317">
        - Selection</span></span><br><span data-ttu-id="0b412-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-318">
        - Settings</span></span><br><span data-ttu-id="0b412-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-319">
        - TableBindings</span></span><br><span data-ttu-id="0b412-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-320">
        - TableCoercion</span></span><br><span data-ttu-id="0b412-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-321">
        - TextBindings</span></span><br><span data-ttu-id="0b412-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="0b412-323">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0b412-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="0b412-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="0b412-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b412-325">Платформа</span><span class="sxs-lookup"><span data-stu-id="0b412-325">Platform</span></span></th>
    <th><span data-ttu-id="0b412-326">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0b412-326">Extension points</span></span></th>
    <th><span data-ttu-id="0b412-327">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0b412-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b412-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b412-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b412-329">Office Online</span></span></td>
    <td> <span data-ttu-id="0b412-330">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-330">- Mail Read</span></span><br><span data-ttu-id="0b412-331">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-331">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b412-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0b412-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0b412-340">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-341">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-342">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-342">- Mail Read</span></span><br><span data-ttu-id="0b412-343">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-343">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0b412-345">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="0b412-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0b412-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b412-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0b412-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0b412-353">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-354">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-355">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-355">- Mail Read</span></span><br><span data-ttu-id="0b412-356">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-356">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0b412-358">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="0b412-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0b412-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b412-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="0b412-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="0b412-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="0b412-366">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-367">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-368">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-368">- Mail Read</span></span><br><span data-ttu-id="0b412-369">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-369">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="0b412-371">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="0b412-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="0b412-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="0b412-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-377">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-378">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-378">- Mail Read</span></span><br><span data-ttu-id="0b412-379">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="0b412-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="0b412-384">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-385">Office 365 для iOS</span><span class="sxs-lookup"><span data-stu-id="0b412-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="0b412-386">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-386">- Mail Read</span></span><br><span data-ttu-id="0b412-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="0b412-393">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-394">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-395">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-395">- Mail Read</span></span><br><span data-ttu-id="0b412-396">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-396">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b412-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0b412-404">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-405">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-406">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-406">- Mail Read</span></span><br><span data-ttu-id="0b412-407">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-407">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b412-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0b412-415">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-416">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-417">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-417">- Mail Read</span></span><br><span data-ttu-id="0b412-418">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="0b412-418">
      - Mail Compose</span></span><br><span data-ttu-id="0b412-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="0b412-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="0b412-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="0b412-426">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-427">Office 365 для Android</span><span class="sxs-lookup"><span data-stu-id="0b412-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="0b412-428">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="0b412-428">- Mail Read</span></span><br><span data-ttu-id="0b412-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="0b412-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="0b412-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="0b412-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="0b412-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="0b412-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="0b412-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="0b412-435">Недоступно</span><span class="sxs-lookup"><span data-stu-id="0b412-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="0b412-436">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0b412-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="0b412-437">Word</span><span class="sxs-lookup"><span data-stu-id="0b412-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b412-438">Платформа</span><span class="sxs-lookup"><span data-stu-id="0b412-438">Platform</span></span></th>
    <th><span data-ttu-id="0b412-439">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0b412-439">Extension points</span></span></th>
    <th><span data-ttu-id="0b412-440">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0b412-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b412-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b412-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b412-442">Office Online</span></span></td>
    <td> <span data-ttu-id="0b412-443">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-443">- TaskPane</span></span><br><span data-ttu-id="0b412-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b412-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b412-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-449">- BindingEvents</span></span><br><span data-ttu-id="0b412-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-451">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-452">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-452">
         - File</span></span><br><span data-ttu-id="0b412-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-454">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-455">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-458">
         - PdfFile</span></span><br><span data-ttu-id="0b412-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-459">
         - Selection</span></span><br><span data-ttu-id="0b412-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-460">
         - Settings</span></span><br><span data-ttu-id="0b412-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-461">
         - TableBindings</span></span><br><span data-ttu-id="0b412-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-462">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-463">
         - TextBindings</span></span><br><span data-ttu-id="0b412-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-464">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-466">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-467">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-467">- TaskPane</span></span><br><span data-ttu-id="0b412-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b412-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b412-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-473">- BindingEvents</span></span><br><span data-ttu-id="0b412-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-474">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-476">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-477">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-477">
         - File</span></span><br><span data-ttu-id="0b412-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-479">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-480">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-483">
         - PdfFile</span></span><br><span data-ttu-id="0b412-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-484">
         - Selection</span></span><br><span data-ttu-id="0b412-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-485">
         - Settings</span></span><br><span data-ttu-id="0b412-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-486">
         - TableBindings</span></span><br><span data-ttu-id="0b412-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-487">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-488">
         - TextBindings</span></span><br><span data-ttu-id="0b412-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-489">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-491">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-492">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-492">- TaskPane</span></span><br><span data-ttu-id="0b412-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b412-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b412-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-498">- BindingEvents</span></span><br><span data-ttu-id="0b412-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-499">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-501">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-502">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-502">
         - File</span></span><br><span data-ttu-id="0b412-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-504">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-505">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-508">
         - PdfFile</span></span><br><span data-ttu-id="0b412-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-509">
         - Selection</span></span><br><span data-ttu-id="0b412-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-510">
         - Settings</span></span><br><span data-ttu-id="0b412-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-511">
         - TableBindings</span></span><br><span data-ttu-id="0b412-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-512">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-513">
         - TextBindings</span></span><br><span data-ttu-id="0b412-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-514">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-516">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-517">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="0b412-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-520">- BindingEvents</span></span><br><span data-ttu-id="0b412-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-521">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-523">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-524">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-524">
         - File</span></span><br><span data-ttu-id="0b412-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-526">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-527">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-530">
         - PdfFile</span></span><br><span data-ttu-id="0b412-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-531">
         - Selection</span></span><br><span data-ttu-id="0b412-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-532">
         - Settings</span></span><br><span data-ttu-id="0b412-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-533">
         - TableBindings</span></span><br><span data-ttu-id="0b412-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-534">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-535">
         - TextBindings</span></span><br><span data-ttu-id="0b412-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-536">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-538">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-539">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b412-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b412-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-541">- BindingEvents</span></span><br><span data-ttu-id="0b412-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-542">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-544">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-545">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-545">
         - File</span></span><br><span data-ttu-id="0b412-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-547">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-548">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-551">
         - PdfFile</span></span><br><span data-ttu-id="0b412-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-552">
         - Selection</span></span><br><span data-ttu-id="0b412-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-553">
         - Settings</span></span><br><span data-ttu-id="0b412-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-554">
         - TableBindings</span></span><br><span data-ttu-id="0b412-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-555">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-556">
         - TextBindings</span></span><br><span data-ttu-id="0b412-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-557">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-559">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="0b412-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="0b412-560">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b412-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b412-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0b412-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0b412-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-565">- BindingEvents</span></span><br><span data-ttu-id="0b412-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-566">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-568">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-569">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-569">
         - File</span></span><br><span data-ttu-id="0b412-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-571">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-572">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-575">
         - PdfFile</span></span><br><span data-ttu-id="0b412-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-576">
         - Selection</span></span><br><span data-ttu-id="0b412-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-577">
         - Settings</span></span><br><span data-ttu-id="0b412-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-578">
         - TableBindings</span></span><br><span data-ttu-id="0b412-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-579">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-580">
         - TextBindings</span></span><br><span data-ttu-id="0b412-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-581">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-583">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-584">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-584">- TaskPane</span></span><br><span data-ttu-id="0b412-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b412-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b412-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0b412-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0b412-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-590">- BindingEvents</span></span><br><span data-ttu-id="0b412-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-591">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-593">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-594">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-594">
         - File</span></span><br><span data-ttu-id="0b412-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-596">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-597">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-600">
         - PdfFile</span></span><br><span data-ttu-id="0b412-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-601">
         - Selection</span></span><br><span data-ttu-id="0b412-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-602">
         - Settings</span></span><br><span data-ttu-id="0b412-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-603">
         - TableBindings</span></span><br><span data-ttu-id="0b412-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-604">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-605">
         - TextBindings</span></span><br><span data-ttu-id="0b412-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-606">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-608">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-609">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-609">- TaskPane</span></span><br><span data-ttu-id="0b412-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="0b412-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="0b412-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="0b412-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="0b412-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="0b412-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="0b412-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-615">- BindingEvents</span></span><br><span data-ttu-id="0b412-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-616">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-618">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-619">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-619">
         - File</span></span><br><span data-ttu-id="0b412-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-621">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-622">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-625">
         - PdfFile</span></span><br><span data-ttu-id="0b412-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-626">
         - Selection</span></span><br><span data-ttu-id="0b412-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-627">
         - Settings</span></span><br><span data-ttu-id="0b412-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-628">
         - TableBindings</span></span><br><span data-ttu-id="0b412-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-629">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-630">
         - TextBindings</span></span><br><span data-ttu-id="0b412-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-631">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-633">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-634">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="0b412-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="0b412-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="0b412-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-637">- BindingEvents</span></span><br><span data-ttu-id="0b412-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-638">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="0b412-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="0b412-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-640">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-641">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="0b412-641">
         - File</span></span><br><span data-ttu-id="0b412-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-643">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-644">
         - MatrixBindings</span></span><br><span data-ttu-id="0b412-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="0b412-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="0b412-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-647">
         - PdfFile</span></span><br><span data-ttu-id="0b412-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-648">
         - Selection</span></span><br><span data-ttu-id="0b412-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="0b412-649">
         - Settings</span></span><br><span data-ttu-id="0b412-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-650">
         - TableBindings</span></span><br><span data-ttu-id="0b412-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-651">
         - TableCoercion</span></span><br><span data-ttu-id="0b412-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="0b412-652">
         - TextBindings</span></span><br><span data-ttu-id="0b412-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-653">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="0b412-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="0b412-655">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0b412-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="0b412-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0b412-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b412-657">Платформа</span><span class="sxs-lookup"><span data-stu-id="0b412-657">Platform</span></span></th>
    <th><span data-ttu-id="0b412-658">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0b412-658">Extension points</span></span></th>
    <th><span data-ttu-id="0b412-659">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0b412-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b412-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b412-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b412-661">Office Online</span></span></td>
    <td> <span data-ttu-id="0b412-662">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-662">- Content</span></span><br><span data-ttu-id="0b412-663">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-663">
         - TaskPane</span></span><br><span data-ttu-id="0b412-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-666">- ActiveView</span></span><br><span data-ttu-id="0b412-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-667">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-668">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-669">
         - File</span></span><br><span data-ttu-id="0b412-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-670">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-671">
         - PdfFile</span></span><br><span data-ttu-id="0b412-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-672">
         - Selection</span></span><br><span data-ttu-id="0b412-673">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-673">
         - Settings</span></span><br><span data-ttu-id="0b412-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-675">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-676">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-676">- Content</span></span><br><span data-ttu-id="0b412-677">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-677">
         - TaskPane</span></span><br><span data-ttu-id="0b412-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-680">- ActiveView</span></span><br><span data-ttu-id="0b412-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-681">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-682">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-683">
         - File</span></span><br><span data-ttu-id="0b412-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-684">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-685">
         - PdfFile</span></span><br><span data-ttu-id="0b412-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-686">
         - Selection</span></span><br><span data-ttu-id="0b412-687">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-687">
         - Settings</span></span><br><span data-ttu-id="0b412-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-689">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-690">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-690">- Content</span></span><br><span data-ttu-id="0b412-691">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-691">
         - TaskPane</span></span><br><span data-ttu-id="0b412-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-694">- ActiveView</span></span><br><span data-ttu-id="0b412-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-695">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-696">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-697">
         - File</span></span><br><span data-ttu-id="0b412-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-698">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-699">
         - PdfFile</span></span><br><span data-ttu-id="0b412-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-700">
         - Selection</span></span><br><span data-ttu-id="0b412-701">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-701">
         - Settings</span></span><br><span data-ttu-id="0b412-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-703">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-704">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-704">- Content</span></span><br><span data-ttu-id="0b412-705">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b412-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b412-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-707">- ActiveView</span></span><br><span data-ttu-id="0b412-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-708">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-709">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-710">
         - File</span></span><br><span data-ttu-id="0b412-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-711">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-712">
         - PdfFile</span></span><br><span data-ttu-id="0b412-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-713">
         - Selection</span></span><br><span data-ttu-id="0b412-714">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-714">
         - Settings</span></span><br><span data-ttu-id="0b412-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-716">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-717">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-717">- Content</span></span><br><span data-ttu-id="0b412-718">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="0b412-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b412-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b412-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-720">- ActiveView</span></span><br><span data-ttu-id="0b412-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-721">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-722">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-723">
         - File</span></span><br><span data-ttu-id="0b412-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-724">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-725">
         - PdfFile</span></span><br><span data-ttu-id="0b412-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-726">
         - Selection</span></span><br><span data-ttu-id="0b412-727">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-727">
         - Settings</span></span><br><span data-ttu-id="0b412-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-729">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="0b412-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="0b412-730">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-730">- Content</span></span><br><span data-ttu-id="0b412-731">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="0b412-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-733">- ActiveView</span></span><br><span data-ttu-id="0b412-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-734">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-735">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-736">
         - File</span></span><br><span data-ttu-id="0b412-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-737">
         - PdfFile</span></span><br><span data-ttu-id="0b412-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-738">
         - Selection</span></span><br><span data-ttu-id="0b412-739">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-739">
         - Settings</span></span><br><span data-ttu-id="0b412-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-740">
         - TextCoercion</span></span><br><span data-ttu-id="0b412-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-742">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-743">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-743">- Content</span></span><br><span data-ttu-id="0b412-744">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-744">
         - TaskPane</span></span><br><span data-ttu-id="0b412-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-747">- ActiveView</span></span><br><span data-ttu-id="0b412-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-748">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-749">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-750">
         - File</span></span><br><span data-ttu-id="0b412-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-751">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-752">
         - PdfFile</span></span><br><span data-ttu-id="0b412-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-753">
         - Selection</span></span><br><span data-ttu-id="0b412-754">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-754">
         - Settings</span></span><br><span data-ttu-id="0b412-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-756">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-757">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-757">- Content</span></span><br><span data-ttu-id="0b412-758">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-758">
         - TaskPane</span></span><br><span data-ttu-id="0b412-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-761">- ActiveView</span></span><br><span data-ttu-id="0b412-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-762">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-763">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-764">
         - File</span></span><br><span data-ttu-id="0b412-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-765">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-766">
         - PdfFile</span></span><br><span data-ttu-id="0b412-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-767">
         - Selection</span></span><br><span data-ttu-id="0b412-768">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-768">
         - Settings</span></span><br><span data-ttu-id="0b412-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-770">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="0b412-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="0b412-771">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-771">- Content</span></span><br><span data-ttu-id="0b412-772">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="0b412-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="0b412-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="0b412-774">- ActiveView</span></span><br><span data-ttu-id="0b412-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="0b412-775">
         - CompressedFile</span></span><br><span data-ttu-id="0b412-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-776">
         - DocumentEvents</span></span><br><span data-ttu-id="0b412-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="0b412-777">
         - File</span></span><br><span data-ttu-id="0b412-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-778">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="0b412-779">
         - PdfFile</span></span><br><span data-ttu-id="0b412-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-780">
         - Selection</span></span><br><span data-ttu-id="0b412-781">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-781">
         - Settings</span></span><br><span data-ttu-id="0b412-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="0b412-783">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="0b412-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="0b412-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="0b412-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b412-785">Платформа</span><span class="sxs-lookup"><span data-stu-id="0b412-785">Platform</span></span></th>
    <th><span data-ttu-id="0b412-786">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0b412-786">Extension points</span></span></th>
    <th><span data-ttu-id="0b412-787">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0b412-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b412-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b412-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="0b412-789">Office Online</span></span></td>
    <td> <span data-ttu-id="0b412-790">- Контент</span><span class="sxs-lookup"><span data-stu-id="0b412-790">- Content</span></span><br><span data-ttu-id="0b412-791">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-791">
         - TaskPane</span></span><br><span data-ttu-id="0b412-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="0b412-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="0b412-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="0b412-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="0b412-795">- DocumentEvents</span></span><br><span data-ttu-id="0b412-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="0b412-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-797">
         - ImageCoercion</span></span><br><span data-ttu-id="0b412-798">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="0b412-798">
         - Settings</span></span><br><span data-ttu-id="0b412-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="0b412-800">Project</span><span class="sxs-lookup"><span data-stu-id="0b412-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="0b412-801">Платформа</span><span class="sxs-lookup"><span data-stu-id="0b412-801">Platform</span></span></th>
    <th><span data-ttu-id="0b412-802">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="0b412-802">Extension points</span></span></th>
    <th><span data-ttu-id="0b412-803">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="0b412-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="0b412-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="0b412-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-805">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-806">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-808">- Selection</span></span><br><span data-ttu-id="0b412-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-810">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-811">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-813">- Selection</span></span><br><span data-ttu-id="0b412-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="0b412-815">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="0b412-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="0b412-816">- Область задач</span><span class="sxs-lookup"><span data-stu-id="0b412-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="0b412-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="0b412-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="0b412-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="0b412-818">- Selection</span></span><br><span data-ttu-id="0b412-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="0b412-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="0b412-820">См. также</span><span class="sxs-lookup"><span data-stu-id="0b412-820">See also</span></span>

- [<span data-ttu-id="0b412-821">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0b412-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="0b412-822">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="0b412-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="0b412-823">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="0b412-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="0b412-824">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="0b412-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
