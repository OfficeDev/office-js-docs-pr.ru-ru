---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 03/15/2019
localization_priority: Priority
ms.openlocfilehash: 4348881c35e4c79975d34406e4668b2693405134
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/16/2019
ms.locfileid: "30654965"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4357b-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="4357b-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4357b-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="4357b-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="4357b-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="4357b-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4357b-p102">Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="4357b-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="4357b-108">Номер сборки для единовременной покупки Office 2019 — 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="4357b-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="4357b-109">Excel</span><span class="sxs-lookup"><span data-stu-id="4357b-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4357b-110">Платформа</span><span class="sxs-lookup"><span data-stu-id="4357b-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4357b-111">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4357b-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4357b-112">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4357b-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4357b-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4357b-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="4357b-114">Office Online</span></span></td>
    <td> <span data-ttu-id="4357b-115">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-115">- TaskPane</span></span><br><span data-ttu-id="4357b-116">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-116">
        - Content</span></span><br><span data-ttu-id="4357b-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="4357b-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4357b-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4357b-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4357b-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4357b-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4357b-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4357b-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4357b-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4357b-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4357b-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4357b-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-127">
        - BindingEvents</span></span><br><span data-ttu-id="4357b-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-128">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-129">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-130">
        - File</span></span><br><span data-ttu-id="4357b-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-131">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-133">
        - Selection</span></span><br><span data-ttu-id="4357b-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-134">
        - Settings</span></span><br><span data-ttu-id="4357b-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-135">
        - TableBindings</span></span><br><span data-ttu-id="4357b-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-136">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-137">
        - TextBindings</span></span><br><span data-ttu-id="4357b-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-139">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-140">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-140">- TaskPane</span></span><br><span data-ttu-id="4357b-141">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-141">
        - Content</span></span><br><span data-ttu-id="4357b-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="4357b-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4357b-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4357b-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4357b-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4357b-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4357b-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4357b-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4357b-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4357b-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4357b-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4357b-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-152">
        - BindingEvents</span></span><br><span data-ttu-id="4357b-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-153">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-154">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-155">
        - File</span></span><br><span data-ttu-id="4357b-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-156">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-158">
        - Selection</span></span><br><span data-ttu-id="4357b-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-159">
        - Settings</span></span><br><span data-ttu-id="4357b-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-160">
        - TableBindings</span></span><br><span data-ttu-id="4357b-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-161">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-162">
        - TextBindings</span></span><br><span data-ttu-id="4357b-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-164">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="4357b-165">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-165">- TaskPane</span></span><br><span data-ttu-id="4357b-166">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-166">
        - Content</span></span><br><span data-ttu-id="4357b-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4357b-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4357b-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4357b-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4357b-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4357b-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4357b-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4357b-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4357b-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4357b-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4357b-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-177">- BindingEvents</span></span><br><span data-ttu-id="4357b-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-178">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-179">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-180">
        - File</span></span><br><span data-ttu-id="4357b-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-181">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-182">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-184">
        - Selection</span></span><br><span data-ttu-id="4357b-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-185">
        - Settings</span></span><br><span data-ttu-id="4357b-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-186">
        - TableBindings</span></span><br><span data-ttu-id="4357b-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-187">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-188">
        - TextBindings</span></span><br><span data-ttu-id="4357b-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-190">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4357b-191">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-191">- TaskPane</span></span><br><span data-ttu-id="4357b-192">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-192">
        - Content</span></span></td>
    <td><span data-ttu-id="4357b-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4357b-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4357b-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-195">- BindingEvents</span></span><br><span data-ttu-id="4357b-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-196">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-197">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-198">
        - File</span></span><br><span data-ttu-id="4357b-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-199">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-200">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-202">
        - Selection</span></span><br><span data-ttu-id="4357b-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-203">
        - Settings</span></span><br><span data-ttu-id="4357b-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-204">
        - TableBindings</span></span><br><span data-ttu-id="4357b-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-205">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-206">
        - TextBindings</span></span><br><span data-ttu-id="4357b-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-208">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4357b-209">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-209">
        - TaskPane</span></span><br><span data-ttu-id="4357b-210">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4357b-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4357b-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="4357b-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-212">
        - BindingEvents</span></span><br><span data-ttu-id="4357b-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-213">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-214">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-215">
        - File</span></span><br><span data-ttu-id="4357b-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-216">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-217">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-219">
        - Selection</span></span><br><span data-ttu-id="4357b-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-220">
        - Settings</span></span><br><span data-ttu-id="4357b-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-221">
        - TableBindings</span></span><br><span data-ttu-id="4357b-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-222">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-223">
        - TextBindings</span></span><br><span data-ttu-id="4357b-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-225">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="4357b-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="4357b-226">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-226">- TaskPane</span></span><br><span data-ttu-id="4357b-227">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-227">
        - Content</span></span></td>
    <td><span data-ttu-id="4357b-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4357b-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4357b-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4357b-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4357b-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4357b-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4357b-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4357b-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4357b-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4357b-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-237">- BindingEvents</span></span><br><span data-ttu-id="4357b-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-238">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-239">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-240">
        - File</span></span><br><span data-ttu-id="4357b-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-241">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-242">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-244">
        - Selection</span></span><br><span data-ttu-id="4357b-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-245">
        - Settings</span></span><br><span data-ttu-id="4357b-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-246">
        - TableBindings</span></span><br><span data-ttu-id="4357b-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-247">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-248">
        - TextBindings</span></span><br><span data-ttu-id="4357b-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-250">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="4357b-251">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-251">- TaskPane</span></span><br><span data-ttu-id="4357b-252">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-252">
        - Content</span></span><br><span data-ttu-id="4357b-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4357b-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4357b-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4357b-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4357b-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4357b-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4357b-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4357b-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4357b-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4357b-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4357b-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-263">- BindingEvents</span></span><br><span data-ttu-id="4357b-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-264">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-265">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-266">
        - File</span></span><br><span data-ttu-id="4357b-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-267">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-268">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-270">
        - PdfFile</span></span><br><span data-ttu-id="4357b-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-271">
        - Selection</span></span><br><span data-ttu-id="4357b-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-272">
        - Settings</span></span><br><span data-ttu-id="4357b-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-273">
        - TableBindings</span></span><br><span data-ttu-id="4357b-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-274">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-275">
        - TextBindings</span></span><br><span data-ttu-id="4357b-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-277">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="4357b-278">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-278">- TaskPane</span></span><br><span data-ttu-id="4357b-279">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-279">
        - Content</span></span><br><span data-ttu-id="4357b-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4357b-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4357b-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4357b-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4357b-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4357b-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4357b-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4357b-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4357b-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4357b-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4357b-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-290">- BindingEvents</span></span><br><span data-ttu-id="4357b-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-291">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-292">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-293">
        - File</span></span><br><span data-ttu-id="4357b-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-294">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-295">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-297">
        - PdfFile</span></span><br><span data-ttu-id="4357b-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-298">
        - Selection</span></span><br><span data-ttu-id="4357b-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-299">
        - Settings</span></span><br><span data-ttu-id="4357b-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-300">
        - TableBindings</span></span><br><span data-ttu-id="4357b-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-301">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-302">
        - TextBindings</span></span><br><span data-ttu-id="4357b-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-304">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4357b-305">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-305">- TaskPane</span></span><br><span data-ttu-id="4357b-306">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-306">
        - Content</span></span></td>
    <td><span data-ttu-id="4357b-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4357b-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4357b-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4357b-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-309">- BindingEvents</span></span><br><span data-ttu-id="4357b-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-310">
        - CompressedFile</span></span><br><span data-ttu-id="4357b-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-311">
        - DocumentEvents</span></span><br><span data-ttu-id="4357b-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="4357b-312">
        - File</span></span><br><span data-ttu-id="4357b-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-313">
        - ImageCoercion</span></span><br><span data-ttu-id="4357b-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-314">
        - MatrixBindings</span></span><br><span data-ttu-id="4357b-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="4357b-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-316">
        - PdfFile</span></span><br><span data-ttu-id="4357b-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-317">
        - Selection</span></span><br><span data-ttu-id="4357b-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-318">
        - Settings</span></span><br><span data-ttu-id="4357b-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-319">
        - TableBindings</span></span><br><span data-ttu-id="4357b-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-320">
        - TableCoercion</span></span><br><span data-ttu-id="4357b-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-321">
        - TextBindings</span></span><br><span data-ttu-id="4357b-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4357b-323">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="4357b-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="4357b-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="4357b-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4357b-325">Платформа</span><span class="sxs-lookup"><span data-stu-id="4357b-325">Platform</span></span></th>
    <th><span data-ttu-id="4357b-326">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4357b-326">Extension points</span></span></th>
    <th><span data-ttu-id="4357b-327">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4357b-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="4357b-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4357b-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="4357b-329">Office Online</span></span></td>
    <td> <span data-ttu-id="4357b-330">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-330">- Mail Read</span></span><br><span data-ttu-id="4357b-331">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-331">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4357b-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4357b-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4357b-340">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-341">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-342">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-342">- Mail Read</span></span><br><span data-ttu-id="4357b-343">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-343">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4357b-345">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="4357b-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4357b-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4357b-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4357b-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4357b-353">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-354">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-355">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-355">- Mail Read</span></span><br><span data-ttu-id="4357b-356">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-356">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4357b-358">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="4357b-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4357b-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4357b-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4357b-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4357b-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4357b-366">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-367">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-368">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-368">- Mail Read</span></span><br><span data-ttu-id="4357b-369">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-369">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4357b-371">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="4357b-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4357b-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4357b-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-377">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-378">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-378">- Mail Read</span></span><br><span data-ttu-id="4357b-379">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="4357b-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4357b-384">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-385">Office 365 для iOS</span><span class="sxs-lookup"><span data-stu-id="4357b-385">Office 365 code snippets for iOS</span></span></td>
    <td> <span data-ttu-id="4357b-386">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-386">- Mail Read</span></span><br><span data-ttu-id="4357b-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4357b-393">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-394">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-395">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-395">- Mail Read</span></span><br><span data-ttu-id="4357b-396">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-396">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4357b-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4357b-404">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-405">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-406">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-406">- Mail Read</span></span><br><span data-ttu-id="4357b-407">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-407">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4357b-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4357b-415">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-416">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-417">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-417">- Mail Read</span></span><br><span data-ttu-id="4357b-418">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="4357b-418">
      - Mail Compose</span></span><br><span data-ttu-id="4357b-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4357b-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4357b-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4357b-426">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-427">Office 365 для Android</span><span class="sxs-lookup"><span data-stu-id="4357b-427">See the Office 365 SDK for Android.</span></span></td>
    <td> <span data-ttu-id="4357b-428">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="4357b-428">- Mail Read</span></span><br><span data-ttu-id="4357b-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4357b-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4357b-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4357b-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4357b-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4357b-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4357b-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4357b-435">Недоступно</span><span class="sxs-lookup"><span data-stu-id="4357b-435">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="4357b-436">Word</span><span class="sxs-lookup"><span data-stu-id="4357b-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4357b-437">Платформа</span><span class="sxs-lookup"><span data-stu-id="4357b-437">Platform</span></span></th>
    <th><span data-ttu-id="4357b-438">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4357b-438">Extension points</span></span></th>
    <th><span data-ttu-id="4357b-439">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4357b-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="4357b-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4357b-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="4357b-441">Office Online</span></span></td>
    <td> <span data-ttu-id="4357b-442">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-442">- TaskPane</span></span><br><span data-ttu-id="4357b-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4357b-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4357b-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-448">- BindingEvents</span></span><br><span data-ttu-id="4357b-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-450">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-451">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-451">
         - File</span></span><br><span data-ttu-id="4357b-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-453">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-454">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-457">
         - PdfFile</span></span><br><span data-ttu-id="4357b-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-458">
         - Selection</span></span><br><span data-ttu-id="4357b-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-459">
         - Settings</span></span><br><span data-ttu-id="4357b-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-460">
         - TableBindings</span></span><br><span data-ttu-id="4357b-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-461">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-462">
         - TextBindings</span></span><br><span data-ttu-id="4357b-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-463">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-465">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-466">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-466">- TaskPane</span></span><br><span data-ttu-id="4357b-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4357b-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4357b-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-472">- BindingEvents</span></span><br><span data-ttu-id="4357b-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-473">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-475">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-476">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-476">
         - File</span></span><br><span data-ttu-id="4357b-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-478">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-479">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-482">
         - PdfFile</span></span><br><span data-ttu-id="4357b-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-483">
         - Selection</span></span><br><span data-ttu-id="4357b-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-484">
         - Settings</span></span><br><span data-ttu-id="4357b-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-485">
         - TableBindings</span></span><br><span data-ttu-id="4357b-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-486">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-487">
         - TextBindings</span></span><br><span data-ttu-id="4357b-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-488">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-490">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-491">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-491">- TaskPane</span></span><br><span data-ttu-id="4357b-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4357b-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4357b-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-497">- BindingEvents</span></span><br><span data-ttu-id="4357b-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-498">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-500">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-501">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-501">
         - File</span></span><br><span data-ttu-id="4357b-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-503">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-504">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-507">
         - PdfFile</span></span><br><span data-ttu-id="4357b-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-508">
         - Selection</span></span><br><span data-ttu-id="4357b-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-509">
         - Settings</span></span><br><span data-ttu-id="4357b-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-510">
         - TableBindings</span></span><br><span data-ttu-id="4357b-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-511">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-512">
         - TextBindings</span></span><br><span data-ttu-id="4357b-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-513">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-515">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-516">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4357b-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4357b-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-519">- BindingEvents</span></span><br><span data-ttu-id="4357b-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-520">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-522">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-523">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-523">
         - File</span></span><br><span data-ttu-id="4357b-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-525">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-526">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-529">
         - PdfFile</span></span><br><span data-ttu-id="4357b-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-530">
         - Selection</span></span><br><span data-ttu-id="4357b-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-531">
         - Settings</span></span><br><span data-ttu-id="4357b-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-532">
         - TableBindings</span></span><br><span data-ttu-id="4357b-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-533">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-534">
         - TextBindings</span></span><br><span data-ttu-id="4357b-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-535">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-537">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-538">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4357b-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4357b-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-540">- BindingEvents</span></span><br><span data-ttu-id="4357b-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-541">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-543">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-544">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-544">
         - File</span></span><br><span data-ttu-id="4357b-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-546">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-547">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-550">
         - PdfFile</span></span><br><span data-ttu-id="4357b-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-551">
         - Selection</span></span><br><span data-ttu-id="4357b-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-552">
         - Settings</span></span><br><span data-ttu-id="4357b-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-553">
         - TableBindings</span></span><br><span data-ttu-id="4357b-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-554">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-555">
         - TextBindings</span></span><br><span data-ttu-id="4357b-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-556">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-558">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="4357b-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4357b-559">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4357b-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4357b-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4357b-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4357b-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-564">- BindingEvents</span></span><br><span data-ttu-id="4357b-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-565">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-567">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-568">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-568">
         - File</span></span><br><span data-ttu-id="4357b-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-570">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-571">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-574">
         - PdfFile</span></span><br><span data-ttu-id="4357b-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-575">
         - Selection</span></span><br><span data-ttu-id="4357b-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-576">
         - Settings</span></span><br><span data-ttu-id="4357b-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-577">
         - TableBindings</span></span><br><span data-ttu-id="4357b-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-578">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-579">
         - TextBindings</span></span><br><span data-ttu-id="4357b-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-580">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-582">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-583">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-583">- TaskPane</span></span><br><span data-ttu-id="4357b-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4357b-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4357b-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4357b-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4357b-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-589">- BindingEvents</span></span><br><span data-ttu-id="4357b-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-590">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-592">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-593">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-593">
         - File</span></span><br><span data-ttu-id="4357b-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-595">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-596">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-599">
         - PdfFile</span></span><br><span data-ttu-id="4357b-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-600">
         - Selection</span></span><br><span data-ttu-id="4357b-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-601">
         - Settings</span></span><br><span data-ttu-id="4357b-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-602">
         - TableBindings</span></span><br><span data-ttu-id="4357b-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-603">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-604">
         - TextBindings</span></span><br><span data-ttu-id="4357b-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-605">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-607">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-608">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-608">- TaskPane</span></span><br><span data-ttu-id="4357b-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4357b-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4357b-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4357b-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4357b-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4357b-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4357b-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-614">- BindingEvents</span></span><br><span data-ttu-id="4357b-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-615">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-617">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-618">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-618">
         - File</span></span><br><span data-ttu-id="4357b-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-620">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-621">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-624">
         - PdfFile</span></span><br><span data-ttu-id="4357b-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-625">
         - Selection</span></span><br><span data-ttu-id="4357b-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-626">
         - Settings</span></span><br><span data-ttu-id="4357b-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-627">
         - TableBindings</span></span><br><span data-ttu-id="4357b-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-628">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-629">
         - TextBindings</span></span><br><span data-ttu-id="4357b-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-630">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-632">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-633">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4357b-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4357b-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4357b-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-636">- BindingEvents</span></span><br><span data-ttu-id="4357b-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-637">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4357b-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="4357b-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-639">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-640">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="4357b-640">
         - File</span></span><br><span data-ttu-id="4357b-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-642">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-643">
         - MatrixBindings</span></span><br><span data-ttu-id="4357b-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="4357b-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4357b-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-646">
         - PdfFile</span></span><br><span data-ttu-id="4357b-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-647">
         - Selection</span></span><br><span data-ttu-id="4357b-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4357b-648">
         - Settings</span></span><br><span data-ttu-id="4357b-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-649">
         - TableBindings</span></span><br><span data-ttu-id="4357b-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-650">
         - TableCoercion</span></span><br><span data-ttu-id="4357b-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4357b-651">
         - TextBindings</span></span><br><span data-ttu-id="4357b-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-652">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4357b-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="4357b-654">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="4357b-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4357b-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4357b-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4357b-656">Платформа</span><span class="sxs-lookup"><span data-stu-id="4357b-656">Platform</span></span></th>
    <th><span data-ttu-id="4357b-657">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4357b-657">Extension points</span></span></th>
    <th><span data-ttu-id="4357b-658">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4357b-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="4357b-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4357b-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="4357b-660">Office Online</span></span></td>
    <td> <span data-ttu-id="4357b-661">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-661">- Content</span></span><br><span data-ttu-id="4357b-662">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-662">
         - TaskPane</span></span><br><span data-ttu-id="4357b-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-665">- ActiveView</span></span><br><span data-ttu-id="4357b-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-666">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-667">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-668">
         - File</span></span><br><span data-ttu-id="4357b-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-669">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-670">
         - PdfFile</span></span><br><span data-ttu-id="4357b-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-671">
         - Selection</span></span><br><span data-ttu-id="4357b-672">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-672">
         - Settings</span></span><br><span data-ttu-id="4357b-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-674">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-675">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-675">- Content</span></span><br><span data-ttu-id="4357b-676">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-676">
         - TaskPane</span></span><br><span data-ttu-id="4357b-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-679">- ActiveView</span></span><br><span data-ttu-id="4357b-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-680">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-681">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-682">
         - File</span></span><br><span data-ttu-id="4357b-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-683">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-684">
         - PdfFile</span></span><br><span data-ttu-id="4357b-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-685">
         - Selection</span></span><br><span data-ttu-id="4357b-686">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-686">
         - Settings</span></span><br><span data-ttu-id="4357b-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-688">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-689">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-689">- Content</span></span><br><span data-ttu-id="4357b-690">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-690">
         - TaskPane</span></span><br><span data-ttu-id="4357b-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-693">- ActiveView</span></span><br><span data-ttu-id="4357b-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-694">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-695">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-696">
         - File</span></span><br><span data-ttu-id="4357b-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-697">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-698">
         - PdfFile</span></span><br><span data-ttu-id="4357b-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-699">
         - Selection</span></span><br><span data-ttu-id="4357b-700">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-700">
         - Settings</span></span><br><span data-ttu-id="4357b-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-702">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-703">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-703">- Content</span></span><br><span data-ttu-id="4357b-704">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4357b-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4357b-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-706">- ActiveView</span></span><br><span data-ttu-id="4357b-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-707">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-708">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-709">
         - File</span></span><br><span data-ttu-id="4357b-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-710">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-711">
         - PdfFile</span></span><br><span data-ttu-id="4357b-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-712">
         - Selection</span></span><br><span data-ttu-id="4357b-713">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-713">
         - Settings</span></span><br><span data-ttu-id="4357b-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-715">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-716">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-716">- Content</span></span><br><span data-ttu-id="4357b-717">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4357b-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4357b-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4357b-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-719">- ActiveView</span></span><br><span data-ttu-id="4357b-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-720">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-721">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-722">
         - File</span></span><br><span data-ttu-id="4357b-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-723">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-724">
         - PdfFile</span></span><br><span data-ttu-id="4357b-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-725">
         - Selection</span></span><br><span data-ttu-id="4357b-726">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-726">
         - Settings</span></span><br><span data-ttu-id="4357b-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-728">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="4357b-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4357b-729">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-729">- Content</span></span><br><span data-ttu-id="4357b-730">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4357b-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-732">- ActiveView</span></span><br><span data-ttu-id="4357b-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-733">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-734">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-735">
         - File</span></span><br><span data-ttu-id="4357b-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-736">
         - PdfFile</span></span><br><span data-ttu-id="4357b-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-737">
         - Selection</span></span><br><span data-ttu-id="4357b-738">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-738">
         - Settings</span></span><br><span data-ttu-id="4357b-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-739">
         - TextCoercion</span></span><br><span data-ttu-id="4357b-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-741">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-742">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-742">- Content</span></span><br><span data-ttu-id="4357b-743">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-743">
         - TaskPane</span></span><br><span data-ttu-id="4357b-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-746">- ActiveView</span></span><br><span data-ttu-id="4357b-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-747">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-748">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-749">
         - File</span></span><br><span data-ttu-id="4357b-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-750">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-751">
         - PdfFile</span></span><br><span data-ttu-id="4357b-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-752">
         - Selection</span></span><br><span data-ttu-id="4357b-753">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-753">
         - Settings</span></span><br><span data-ttu-id="4357b-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-755">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-756">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-756">- Content</span></span><br><span data-ttu-id="4357b-757">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-757">
         - TaskPane</span></span><br><span data-ttu-id="4357b-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-760">- ActiveView</span></span><br><span data-ttu-id="4357b-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-761">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-762">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-763">
         - File</span></span><br><span data-ttu-id="4357b-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-764">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-765">
         - PdfFile</span></span><br><span data-ttu-id="4357b-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-766">
         - Selection</span></span><br><span data-ttu-id="4357b-767">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-767">
         - Settings</span></span><br><span data-ttu-id="4357b-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-769">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="4357b-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4357b-770">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-770">- Content</span></span><br><span data-ttu-id="4357b-771">
         - Область задач/td></span><span class="sxs-lookup"><span data-stu-id="4357b-771">
         - TaskPane/td></span></span> <td> <span data-ttu-id="4357b-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4357b-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4357b-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4357b-773">- ActiveView</span></span><br><span data-ttu-id="4357b-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4357b-774">
         - CompressedFile</span></span><br><span data-ttu-id="4357b-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-775">
         - DocumentEvents</span></span><br><span data-ttu-id="4357b-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="4357b-776">
         - File</span></span><br><span data-ttu-id="4357b-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-777">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4357b-778">
         - PdfFile</span></span><br><span data-ttu-id="4357b-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-779">
         - Selection</span></span><br><span data-ttu-id="4357b-780">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-780">
         - Settings</span></span><br><span data-ttu-id="4357b-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4357b-782">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="4357b-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4357b-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="4357b-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4357b-784">Платформа</span><span class="sxs-lookup"><span data-stu-id="4357b-784">Platform</span></span></th>
    <th><span data-ttu-id="4357b-785">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4357b-785">Extension points</span></span></th>
    <th><span data-ttu-id="4357b-786">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4357b-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="4357b-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4357b-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="4357b-788">Office Online</span></span></td>
    <td> <span data-ttu-id="4357b-789">- Контент</span><span class="sxs-lookup"><span data-stu-id="4357b-789">- Content</span></span><br><span data-ttu-id="4357b-790">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-790">
         - TaskPane</span></span><br><span data-ttu-id="4357b-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="4357b-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4357b-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4357b-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4357b-794">- DocumentEvents</span></span><br><span data-ttu-id="4357b-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="4357b-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-796">
         - ImageCoercion</span></span><br><span data-ttu-id="4357b-797">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="4357b-797">
         - Settings</span></span><br><span data-ttu-id="4357b-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4357b-799">Project</span><span class="sxs-lookup"><span data-stu-id="4357b-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4357b-800">Платформа</span><span class="sxs-lookup"><span data-stu-id="4357b-800">Platform</span></span></th>
    <th><span data-ttu-id="4357b-801">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="4357b-801">Extension points</span></span></th>
    <th><span data-ttu-id="4357b-802">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="4357b-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="4357b-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="4357b-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-804">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-805">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-807">- Selection</span></span><br><span data-ttu-id="4357b-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-809">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-810">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-812">- Selection</span></span><br><span data-ttu-id="4357b-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4357b-814">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="4357b-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4357b-815">- Область задач</span><span class="sxs-lookup"><span data-stu-id="4357b-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4357b-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4357b-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4357b-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="4357b-817">- Selection</span></span><br><span data-ttu-id="4357b-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4357b-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4357b-819">См. также</span><span class="sxs-lookup"><span data-stu-id="4357b-819">See also</span></span>

- [<span data-ttu-id="4357b-820">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="4357b-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4357b-821">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="4357b-821">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4357b-822">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="4357b-822">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4357b-823">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="4357b-823">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
