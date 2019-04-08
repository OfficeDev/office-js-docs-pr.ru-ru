---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, Word, Outlook, PowerPoint, OneNote и Project.
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477595"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="01242-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="01242-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="01242-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="01242-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="01242-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="01242-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="01242-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="01242-106">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="01242-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="01242-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="01242-108">Excel</span><span class="sxs-lookup"><span data-stu-id="01242-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="01242-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="01242-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="01242-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="01242-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="01242-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="01242-111">API requirement sets</span></span></th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="01242-112">Общие API</span><span class="sxs-lookup"><span data-stu-id="01242-112">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="01242-113">Office Online</span></span></td>
    <td> - <span data-ttu-id="01242-114">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-114">TaskPane</span></span><br>
        - <span data-ttu-id="01242-115">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-115">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-116">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-116">Add-in Commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-117">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-117">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-118">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-118">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-119">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-119">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-120">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-120">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-121">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-122">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-122">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-123">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-123">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-124">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="01242-124">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-125">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-125">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="01242-126">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-126">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-127">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-127">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-128">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-128">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-129">File</span><span class="sxs-lookup"><span data-stu-id="01242-129">File</span></span><br>
        - <span data-ttu-id="01242-130">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-130">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-131">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-131">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-132">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-132">Selection</span></span><br>
        - <span data-ttu-id="01242-133">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-133">Settings</span></span><br>
        - <span data-ttu-id="01242-134">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-134">TableBindings</span></span><br>
        - <span data-ttu-id="01242-135">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-135">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-136">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-136">TextBindings</span></span><br>
        - <span data-ttu-id="01242-137">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-137">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-138">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-138">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-139">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-139">TaskPane</span></span><br>
        - <span data-ttu-id="01242-140">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-140">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-141">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-141">Add-in Commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-142">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-142">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-143">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-143">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-144">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-144">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-145">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-145">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-146">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-146">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-147">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-147">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-148">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-148">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-149">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="01242-149">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-150">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-150">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="01242-151">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-151">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-152">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-152">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-153">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-153">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-154">File</span><span class="sxs-lookup"><span data-stu-id="01242-154">File</span></span><br>
        - <span data-ttu-id="01242-155">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-155">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-156">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-156">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-157">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-157">Selection</span></span><br>
        - <span data-ttu-id="01242-158">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-158">Settings</span></span><br>
        - <span data-ttu-id="01242-159">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-159">TableBindings</span></span><br>
        - <span data-ttu-id="01242-160">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-160">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-161">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-161">TextBindings</span></span><br>
        - <span data-ttu-id="01242-162">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-162">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-163">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-163">Office 2019 for Windows</span></span></td>
    <td>- <span data-ttu-id="01242-164">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-164">TaskPane</span></span><br>
        - <span data-ttu-id="01242-165">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-165">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-166">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-166">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-167">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-167">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-168">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-168">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-169">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-169">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-170">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-170">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-171">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-171">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-172">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-172">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-173">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-173">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-174">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="01242-174">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-175">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-175">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="01242-176">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-176">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-177">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-177">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-178">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-178">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-179">File</span><span class="sxs-lookup"><span data-stu-id="01242-179">File</span></span><br>
        - <span data-ttu-id="01242-180">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-180">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-181">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-181">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-182">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-182">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-183">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-183">Selection</span></span><br>
        - <span data-ttu-id="01242-184">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-184">Settings</span></span><br>
        - <span data-ttu-id="01242-185">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-185">TableBindings</span></span><br>
        - <span data-ttu-id="01242-186">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-186">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-187">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-187">TextBindings</span></span><br>
        - <span data-ttu-id="01242-188">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-188">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-189">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-189">Office 2016 for Windows</span></span></td>
    <td>- <span data-ttu-id="01242-190">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-190">TaskPane</span></span><br>
        - <span data-ttu-id="01242-191">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-191">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-192">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-192">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-193">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-193">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="01242-194">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-194">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-195">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-195">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-196">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-196">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-197">File</span><span class="sxs-lookup"><span data-stu-id="01242-197">File</span></span><br>
        - <span data-ttu-id="01242-198">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-198">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-199">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-199">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-200">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-200">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-201">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-201">Selection</span></span><br>
        - <span data-ttu-id="01242-202">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-202">Settings</span></span><br>
        - <span data-ttu-id="01242-203">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-203">TableBindings</span></span><br>
        - <span data-ttu-id="01242-204">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-204">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-205">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-205">TextBindings</span></span><br>
        - <span data-ttu-id="01242-206">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-206">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-207">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-207">Office 2013 for Windows</span></span></td>
    <td>
        - <span data-ttu-id="01242-208">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-208">TaskPane</span></span><br>
        - <span data-ttu-id="01242-209">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-209">Content</span></span></td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-210">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-210">DialogApi 1.1</span></span></a>*</td>
    <td>
        - <span data-ttu-id="01242-211">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-211">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-212">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-212">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-213">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-213">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-214">File</span><span class="sxs-lookup"><span data-stu-id="01242-214">File</span></span><br>
        - <span data-ttu-id="01242-215">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-215">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-216">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-216">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-217">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-217">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-218">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-218">Selection</span></span><br>
        - <span data-ttu-id="01242-219">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-219">Settings</span></span><br>
        - <span data-ttu-id="01242-220">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-220">TableBindings</span></span><br>
        - <span data-ttu-id="01242-221">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-221">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-222">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-222">TextBindings</span></span><br>
        - <span data-ttu-id="01242-223">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-223">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-224">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="01242-224">Office 365 for iPad</span></span></td>
    <td>- <span data-ttu-id="01242-225">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-225">TaskPane</span></span><br>
        - <span data-ttu-id="01242-226">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-226">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-227">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-227">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-228">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-228">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-229">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-229">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-230">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-230">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-231">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-231">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-232">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-232">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-233">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-233">ExcelApi 1.7</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-234">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="01242-234">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-235">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-235">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="01242-236">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-236">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-237">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-237">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-238">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-238">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-239">File</span><span class="sxs-lookup"><span data-stu-id="01242-239">File</span></span><br>
        - <span data-ttu-id="01242-240">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-240">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-241">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-241">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-242">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-242">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-243">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-243">Selection</span></span><br>
        - <span data-ttu-id="01242-244">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-244">Settings</span></span><br>
        - <span data-ttu-id="01242-245">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-245">TableBindings</span></span><br>
        - <span data-ttu-id="01242-246">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-246">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-247">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-247">TextBindings</span></span><br>
        - <span data-ttu-id="01242-248">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-248">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-249">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-249">Office 365 for Mac</span></span></td>
    <td>- <span data-ttu-id="01242-250">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-250">TaskPane</span></span><br>
        - <span data-ttu-id="01242-251">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-251">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-252">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-252">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-253">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-253">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-254">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-254">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-255">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-255">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-256">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-256">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-257">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-257">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-258">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-258">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-259">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-259">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-260">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="01242-260">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-261">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-261">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="01242-262">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-262">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-263">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-263">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-264">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-264">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-265">File</span><span class="sxs-lookup"><span data-stu-id="01242-265">File</span></span><br>
        - <span data-ttu-id="01242-266">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-266">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-267">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-267">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-268">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-268">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-269">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-269">PdfFile</span></span><br>
        - <span data-ttu-id="01242-270">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-270">Selection</span></span><br>
        - <span data-ttu-id="01242-271">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-271">Settings</span></span><br>
        - <span data-ttu-id="01242-272">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-272">TableBindings</span></span><br>
        - <span data-ttu-id="01242-273">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-273">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-274">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-274">TextBindings</span></span><br>
        - <span data-ttu-id="01242-275">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-275">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-276">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-276">Office 2019 for Mac</span></span></td>
    <td>- <span data-ttu-id="01242-277">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-277">TaskPane</span></span><br>
        - <span data-ttu-id="01242-278">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-278">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-279">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-279">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-280">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-280">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-281">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-281">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-282">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-282">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-283">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-283">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-284">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-284">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-285">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-285">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-286">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-286">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-287">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="01242-287">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-288">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-288">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="01242-289">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-289">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-290">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-290">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-291">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-291">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-292">File</span><span class="sxs-lookup"><span data-stu-id="01242-292">File</span></span><br>
        - <span data-ttu-id="01242-293">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-293">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-294">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-294">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-295">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-295">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-296">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-296">PdfFile</span></span><br>
        - <span data-ttu-id="01242-297">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-297">Selection</span></span><br>
        - <span data-ttu-id="01242-298">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-298">Settings</span></span><br>
        - <span data-ttu-id="01242-299">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-299">TableBindings</span></span><br>
        - <span data-ttu-id="01242-300">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-300">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-301">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-301">TextBindings</span></span><br>
        - <span data-ttu-id="01242-302">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-302">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-303">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-303">Office 2016 for Mac</span></span></td>
    <td>- <span data-ttu-id="01242-304">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-304">TaskPane</span></span><br>
        - <span data-ttu-id="01242-305">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-305">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="01242-306">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-306">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-307">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-307">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="01242-308">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-308">BindingEvents</span></span><br>
        - <span data-ttu-id="01242-309">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-309">CompressedFile</span></span><br>
        - <span data-ttu-id="01242-310">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-310">DocumentEvents</span></span><br>
        - <span data-ttu-id="01242-311">File</span><span class="sxs-lookup"><span data-stu-id="01242-311">File</span></span><br>
        - <span data-ttu-id="01242-312">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-312">ImageCoercion</span></span><br>
        - <span data-ttu-id="01242-313">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-313">MatrixBindings</span></span><br>
        - <span data-ttu-id="01242-314">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-314">MatrixCoercion</span></span><br>
        - <span data-ttu-id="01242-315">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-315">PdfFile</span></span><br>
        - <span data-ttu-id="01242-316">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-316">Selection</span></span><br>
        - <span data-ttu-id="01242-317">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-317">Settings</span></span><br>
        - <span data-ttu-id="01242-318">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-318">TableBindings</span></span><br>
        - <span data-ttu-id="01242-319">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-319">TableCoercion</span></span><br>
        - <span data-ttu-id="01242-320">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-320">TextBindings</span></span><br>
        - <span data-ttu-id="01242-321">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-321">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="01242-322">&ast; — Добавлены обновления после выпуска.</span><span class="sxs-lookup"><span data-stu-id="01242-322">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="outlook"></a><span data-ttu-id="01242-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="01242-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="01242-324">Платформа</span><span class="sxs-lookup"><span data-stu-id="01242-324">Platform</span></span></th>
    <th><span data-ttu-id="01242-325">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="01242-325">Extension points</span></span></th>
    <th><span data-ttu-id="01242-326">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="01242-326">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="01242-327">Общие API</span><span class="sxs-lookup"><span data-stu-id="01242-327">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="01242-328">Office Online</span></span></td>
    <td> - <span data-ttu-id="01242-329">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-329">Mail Read</span></span><br>
      - <span data-ttu-id="01242-330">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-330">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-331">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-331">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-332">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-332">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-333">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-333">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-334">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-334">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-335">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-335">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-336">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-336">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="01242-337">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-337">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="01242-338">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-338">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="01242-339">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-340">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-340">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-341">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-341">Mail Read</span></span><br>
      - <span data-ttu-id="01242-342">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-342">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-343">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-343">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="01242-344">Модули</span><span class="sxs-lookup"><span data-stu-id="01242-344">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-345">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-345">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-346">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-346">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-347">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-347">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-348">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-348">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-349">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-349">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="01242-350">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-350">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="01242-351">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-351">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="01242-352">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-353">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-353">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-354">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-354">Mail Read</span></span><br>
      - <span data-ttu-id="01242-355">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-355">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-356">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-356">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="01242-357">Модули</span><span class="sxs-lookup"><span data-stu-id="01242-357">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-358">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-358">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-359">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-359">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-360">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-360">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-361">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-361">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-362">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-362">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="01242-363">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-363">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="01242-364">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="01242-364">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="01242-365">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-366">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-366">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-367">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-367">Mail Read</span></span><br>
      - <span data-ttu-id="01242-368">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-368">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-369">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-369">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="01242-370">Модули</span><span class="sxs-lookup"><span data-stu-id="01242-370">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-371">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-371">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-372">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-372">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-373">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-373">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-374">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-374">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="01242-375">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-376">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-376">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-377">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-377">Mail Read</span></span><br>
      - <span data-ttu-id="01242-378">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-378">Mail Compose</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-379">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-379">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-380">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-380">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-381">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-381">Mailbox 1.3</span></span></a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-382">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-382">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="01242-383">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-384">Office 365 для iOS</span><span class="sxs-lookup"><span data-stu-id="01242-384">Office 365 for iOS</span></span></td>
    <td> - <span data-ttu-id="01242-385">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-385">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-386">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-386">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-387">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-387">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-388">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-388">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-389">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-389">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-390">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-390">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-391">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-391">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="01242-392">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-393">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-393">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-394">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-394">Mail Read</span></span><br>
      - <span data-ttu-id="01242-395">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-395">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-396">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-396">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-397">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-397">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-398">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-398">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-399">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-399">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-400">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-400">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-401">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-401">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="01242-402">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-402">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="01242-403">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-404">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-404">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-405">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-405">Mail Read</span></span><br>
      - <span data-ttu-id="01242-406">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-406">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-407">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-407">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-408">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-408">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-409">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-409">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-410">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-410">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-411">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-411">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-412">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-412">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="01242-413">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-413">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="01242-414">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-415">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-415">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-416">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-416">Mail Read</span></span><br>
      - <span data-ttu-id="01242-417">Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="01242-417">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-418">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-418">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-419">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-419">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-420">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-420">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-421">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-421">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-422">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-422">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-423">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-423">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="01242-424">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="01242-424">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="01242-425">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-426">Office 365 для Android</span><span class="sxs-lookup"><span data-stu-id="01242-426">Office 365 for Android</span></span></td>
    <td> - <span data-ttu-id="01242-427">Чтение почты</span><span class="sxs-lookup"><span data-stu-id="01242-427">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-428">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-428">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="01242-429">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-429">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="01242-430">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-430">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="01242-431">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-431">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="01242-432">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="01242-432">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="01242-433">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="01242-433">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="01242-434">Недоступно</span><span class="sxs-lookup"><span data-stu-id="01242-434">Not available</span></span></td>
  </tr>
</table>

*<span data-ttu-id="01242-435">&ast; — Добавлены обновления после выпуска.</span><span class="sxs-lookup"><span data-stu-id="01242-435">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="word"></a><span data-ttu-id="01242-436">Word</span><span class="sxs-lookup"><span data-stu-id="01242-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="01242-437">Платформа</span><span class="sxs-lookup"><span data-stu-id="01242-437">Platform</span></span></th>
    <th><span data-ttu-id="01242-438">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="01242-438">Extension points</span></span></th>
    <th><span data-ttu-id="01242-439">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="01242-439">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="01242-440">Общие API</span><span class="sxs-lookup"><span data-stu-id="01242-440">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="01242-441">Office Online</span></span></td>
    <td> - <span data-ttu-id="01242-442">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-442">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-443">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-443">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-444">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-444">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-445">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-445">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-446">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-446">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-447">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-447">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-448">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-448">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-449">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-449">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-450">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-450">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-451">File</span><span class="sxs-lookup"><span data-stu-id="01242-451">File</span></span><br>
         - <span data-ttu-id="01242-452">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-452">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-453">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-453">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-454">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-454">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-455">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-455">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-456">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-456">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-457">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-457">PdfFile</span></span><br>
         - <span data-ttu-id="01242-458">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-458">Selection</span></span><br>
         - <span data-ttu-id="01242-459">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-459">Settings</span></span><br>
         - <span data-ttu-id="01242-460">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-460">TableBindings</span></span><br>
         - <span data-ttu-id="01242-461">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-461">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-462">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-462">TextBindings</span></span><br>
         - <span data-ttu-id="01242-463">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-463">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-464">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-464">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-465">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-465">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-466">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-466">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-467">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-467">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-468">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-468">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-469">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-469">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-470">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-470">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-471">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-471">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-472">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-472">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-473">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-473">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-474">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-474">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-475">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-475">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-476">File</span><span class="sxs-lookup"><span data-stu-id="01242-476">File</span></span><br>
         - <span data-ttu-id="01242-477">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-477">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-478">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-478">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-479">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-479">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-480">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-480">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-481">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-481">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-482">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-482">PdfFile</span></span><br>
         - <span data-ttu-id="01242-483">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-483">Selection</span></span><br>
         - <span data-ttu-id="01242-484">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-484">Settings</span></span><br>
         - <span data-ttu-id="01242-485">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-485">TableBindings</span></span><br>
         - <span data-ttu-id="01242-486">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-486">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-487">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-487">TextBindings</span></span><br>
         - <span data-ttu-id="01242-488">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-488">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-489">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-489">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-490">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-490">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-491">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-491">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-492">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-492">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-493">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-493">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-494">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-494">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-495">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-495">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-496">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-496">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-497">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-497">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-498">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-498">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-499">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-499">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-500">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-500">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-501">File</span><span class="sxs-lookup"><span data-stu-id="01242-501">File</span></span><br>
         - <span data-ttu-id="01242-502">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-502">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-503">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-503">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-504">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-504">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-505">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-505">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-506">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-506">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-507">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-507">PdfFile</span></span><br>
         - <span data-ttu-id="01242-508">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-508">Selection</span></span><br>
         - <span data-ttu-id="01242-509">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-509">Settings</span></span><br>
         - <span data-ttu-id="01242-510">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-510">TableBindings</span></span><br>
         - <span data-ttu-id="01242-511">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-511">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-512">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-512">TextBindings</span></span><br>
         - <span data-ttu-id="01242-513">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-513">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-514">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-514">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-515">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-515">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-516">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-516">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-517">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-517">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-518">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-518">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="01242-519">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-519">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-520">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-520">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-521">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-521">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-522">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-522">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-523">File</span><span class="sxs-lookup"><span data-stu-id="01242-523">File</span></span><br>
         - <span data-ttu-id="01242-524">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-524">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-525">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-525">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-526">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-526">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-527">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-527">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-528">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-528">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-529">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-529">PdfFile</span></span><br>
         - <span data-ttu-id="01242-530">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-530">Selection</span></span><br>
         - <span data-ttu-id="01242-531">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-531">Settings</span></span><br>
         - <span data-ttu-id="01242-532">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-532">TableBindings</span></span><br>
         - <span data-ttu-id="01242-533">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-533">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-534">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-534">TextBindings</span></span><br>
         - <span data-ttu-id="01242-535">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-535">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-536">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-536">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-537">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-537">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-538">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-538">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-539">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-539">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="01242-540">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-540">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-541">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-541">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-542">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-542">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-543">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-543">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-544">File</span><span class="sxs-lookup"><span data-stu-id="01242-544">File</span></span><br>
         - <span data-ttu-id="01242-545">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-545">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-546">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-546">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-547">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-547">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-548">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-548">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-549">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-549">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-550">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-550">PdfFile</span></span><br>
         - <span data-ttu-id="01242-551">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-551">Selection</span></span><br>
         - <span data-ttu-id="01242-552">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-552">Settings</span></span><br>
         - <span data-ttu-id="01242-553">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-553">TableBindings</span></span><br>
         - <span data-ttu-id="01242-554">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-554">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-555">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-555">TextBindings</span></span><br>
         - <span data-ttu-id="01242-556">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-556">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-557">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-557">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-558">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="01242-558">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="01242-559">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-559">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-560">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-560">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-561">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-561">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-562">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-562">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-563">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-563">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="01242-564">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-564">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-565">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-565">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-566">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-566">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-567">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-567">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-568">File</span><span class="sxs-lookup"><span data-stu-id="01242-568">File</span></span><br>
         - <span data-ttu-id="01242-569">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-569">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-570">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-570">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-571">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-571">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-572">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-572">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-573">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-573">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-574">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-574">PdfFile</span></span><br>
         - <span data-ttu-id="01242-575">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-575">Selection</span></span><br>
         - <span data-ttu-id="01242-576">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-576">Settings</span></span><br>
         - <span data-ttu-id="01242-577">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-577">TableBindings</span></span><br>
         - <span data-ttu-id="01242-578">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-578">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-579">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-579">TextBindings</span></span><br>
         - <span data-ttu-id="01242-580">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-580">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-581">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-581">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-582">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-582">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-583">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-583">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-584">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-584">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-585">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-585">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-586">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-586">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-587">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-587">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-588">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-588">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="01242-589">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-589">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-590">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-590">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-591">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-591">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-592">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-592">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-593">File</span><span class="sxs-lookup"><span data-stu-id="01242-593">File</span></span><br>
         - <span data-ttu-id="01242-594">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-594">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-595">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-595">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-596">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-596">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-597">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-597">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-598">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-598">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-599">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-599">PdfFile</span></span><br>
         - <span data-ttu-id="01242-600">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-600">Selection</span></span><br>
         - <span data-ttu-id="01242-601">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-601">Settings</span></span><br>
         - <span data-ttu-id="01242-602">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-602">TableBindings</span></span><br>
         - <span data-ttu-id="01242-603">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-603">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-604">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-604">TextBindings</span></span><br>
         - <span data-ttu-id="01242-605">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-605">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-606">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-606">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-607">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-607">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-608">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-608">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-609">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-609">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-610">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-610">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-611">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="01242-611">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-612">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="01242-612">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-613">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-613">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="01242-614">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-614">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-615">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-615">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-616">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-616">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-617">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-617">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-618">File</span><span class="sxs-lookup"><span data-stu-id="01242-618">File</span></span><br>
         - <span data-ttu-id="01242-619">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-619">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-620">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-620">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-621">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-621">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-622">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-622">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-623">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-623">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-624">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-624">PdfFile</span></span><br>
         - <span data-ttu-id="01242-625">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-625">Selection</span></span><br>
         - <span data-ttu-id="01242-626">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-626">Settings</span></span><br>
         - <span data-ttu-id="01242-627">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-627">TableBindings</span></span><br>
         - <span data-ttu-id="01242-628">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-628">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-629">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-629">TextBindings</span></span><br>
         - <span data-ttu-id="01242-630">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-630">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-631">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-631">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-632">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-632">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-633">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-633">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="01242-634">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-634">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-635">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-635">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="01242-636">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="01242-636">BindingEvents</span></span><br>
         - <span data-ttu-id="01242-637">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-637">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-638">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="01242-638">CustomXmlParts</span></span><br>
         - <span data-ttu-id="01242-639">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-639">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-640">File</span><span class="sxs-lookup"><span data-stu-id="01242-640">File</span></span><br>
         - <span data-ttu-id="01242-641">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-641">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-642">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-642">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-643">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="01242-643">MatrixBindings</span></span><br>
         - <span data-ttu-id="01242-644">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-644">MatrixCoercion</span></span><br>
         - <span data-ttu-id="01242-645">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-645">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="01242-646">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-646">PdfFile</span></span><br>
         - <span data-ttu-id="01242-647">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-647">Selection</span></span><br>
         - <span data-ttu-id="01242-648">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-648">Settings</span></span><br>
         - <span data-ttu-id="01242-649">TableBindings</span><span class="sxs-lookup"><span data-stu-id="01242-649">TableBindings</span></span><br>
         - <span data-ttu-id="01242-650">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-650">TableCoercion</span></span><br>
         - <span data-ttu-id="01242-651">TextBindings</span><span class="sxs-lookup"><span data-stu-id="01242-651">TextBindings</span></span><br>
         - <span data-ttu-id="01242-652">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-652">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-653">TextFile</span><span class="sxs-lookup"><span data-stu-id="01242-653">TextFile</span></span> </td>
  </tr>
</table>

*<span data-ttu-id="01242-654">&ast; — Добавлены обновления после выпуска.</span><span class="sxs-lookup"><span data-stu-id="01242-654">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="powerpoint"></a><span data-ttu-id="01242-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="01242-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="01242-656">Платформа</span><span class="sxs-lookup"><span data-stu-id="01242-656">Platform</span></span></th>
    <th><span data-ttu-id="01242-657">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="01242-657">Extension points</span></span></th>
    <th><span data-ttu-id="01242-658">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="01242-658">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="01242-659">Общие API</span><span class="sxs-lookup"><span data-stu-id="01242-659">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="01242-660">Office Online</span></span></td>
    <td> - <span data-ttu-id="01242-661">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-661">Content</span></span><br>
         - <span data-ttu-id="01242-662">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-662">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-663">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-663">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-664">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-664">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-665">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-665">ActiveView</span></span><br>
         - <span data-ttu-id="01242-666">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-666">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-667">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-667">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-668">File</span><span class="sxs-lookup"><span data-stu-id="01242-668">File</span></span><br>
         - <span data-ttu-id="01242-669">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-669">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-670">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-670">PdfFile</span></span><br>
         - <span data-ttu-id="01242-671">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-671">Selection</span></span><br>
         - <span data-ttu-id="01242-672">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-672">Settings</span></span><br>
         - <span data-ttu-id="01242-673">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-673">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-674">Office 365 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-674">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-675">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-675">Content</span></span><br>
         - <span data-ttu-id="01242-676">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-676">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-677">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-677">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-678">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-678">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-679">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-679">ActiveView</span></span><br>
         - <span data-ttu-id="01242-680">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-680">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-681">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-681">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-682">File</span><span class="sxs-lookup"><span data-stu-id="01242-682">File</span></span><br>
         - <span data-ttu-id="01242-683">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-683">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-684">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-684">PdfFile</span></span><br>
         - <span data-ttu-id="01242-685">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-685">Selection</span></span><br>
         - <span data-ttu-id="01242-686">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-686">Settings</span></span><br>
         - <span data-ttu-id="01242-687">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-687">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-688">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-688">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-689">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-689">Content</span></span><br>
         - <span data-ttu-id="01242-690">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-690">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-691">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-691">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-692">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-692">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-693">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-693">ActiveView</span></span><br>
         - <span data-ttu-id="01242-694">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-694">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-695">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-695">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-696">File</span><span class="sxs-lookup"><span data-stu-id="01242-696">File</span></span><br>
         - <span data-ttu-id="01242-697">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-697">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-698">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-698">PdfFile</span></span><br>
         - <span data-ttu-id="01242-699">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-699">Selection</span></span><br>
         - <span data-ttu-id="01242-700">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-700">Settings</span></span><br>
         - <span data-ttu-id="01242-701">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-701">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-702">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-702">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-703">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-703">Content</span></span><br>
         - <span data-ttu-id="01242-704">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-704">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-705">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-705">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="01242-706">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-706">ActiveView</span></span><br>
         - <span data-ttu-id="01242-707">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-707">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-708">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-708">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-709">File</span><span class="sxs-lookup"><span data-stu-id="01242-709">File</span></span><br>
         - <span data-ttu-id="01242-710">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-710">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-711">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-711">PdfFile</span></span><br>
         - <span data-ttu-id="01242-712">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-712">Selection</span></span><br>
         - <span data-ttu-id="01242-713">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-713">Settings</span></span><br>
         - <span data-ttu-id="01242-714">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-714">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-715">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-715">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-716">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-716">Content</span></span><br>
         - <span data-ttu-id="01242-717">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-717">TaskPane</span></span><br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-718">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-718">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="01242-719">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-719">ActiveView</span></span><br>
         - <span data-ttu-id="01242-720">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-720">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-721">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-721">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-722">File</span><span class="sxs-lookup"><span data-stu-id="01242-722">File</span></span><br>
         - <span data-ttu-id="01242-723">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-723">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-724">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-724">PdfFile</span></span><br>
         - <span data-ttu-id="01242-725">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-725">Selection</span></span><br>
         - <span data-ttu-id="01242-726">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-726">Settings</span></span><br>
         - <span data-ttu-id="01242-727">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-727">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-728">Office 365 для iPad</span><span class="sxs-lookup"><span data-stu-id="01242-728">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="01242-729">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-729">Content</span></span><br>
         - <span data-ttu-id="01242-730">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-730">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-731">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-731">DialogApi 1.1</span></span></a></td>
     <td> - <span data-ttu-id="01242-732">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-732">ActiveView</span></span><br>
         - <span data-ttu-id="01242-733">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-733">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-734">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-734">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-735">File</span><span class="sxs-lookup"><span data-stu-id="01242-735">File</span></span><br>
         - <span data-ttu-id="01242-736">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-736">PdfFile</span></span><br>
         - <span data-ttu-id="01242-737">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-737">Selection</span></span><br>
         - <span data-ttu-id="01242-738">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-738">Settings</span></span><br>
         - <span data-ttu-id="01242-739">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-739">TextCoercion</span></span><br>
         - <span data-ttu-id="01242-740">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-740">ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-741">Office 365 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-741">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-742">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-742">Content</span></span><br>
         - <span data-ttu-id="01242-743">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-743">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-744">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-744">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-745">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-745">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-746">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-746">ActiveView</span></span><br>
         - <span data-ttu-id="01242-747">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-747">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-748">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-748">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-749">File</span><span class="sxs-lookup"><span data-stu-id="01242-749">File</span></span><br>
         - <span data-ttu-id="01242-750">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-750">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-751">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-751">PdfFile</span></span><br>
         - <span data-ttu-id="01242-752">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-752">Selection</span></span><br>
         - <span data-ttu-id="01242-753">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-753">Settings</span></span><br>
         - <span data-ttu-id="01242-754">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-754">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-755">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-755">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-756">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-756">Content</span></span><br>
         - <span data-ttu-id="01242-757">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-757">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-758">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-758">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-759">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-759">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-760">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-760">ActiveView</span></span><br>
         - <span data-ttu-id="01242-761">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-761">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-762">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-762">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-763">File</span><span class="sxs-lookup"><span data-stu-id="01242-763">File</span></span><br>
         - <span data-ttu-id="01242-764">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-764">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-765">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-765">PdfFile</span></span><br>
         - <span data-ttu-id="01242-766">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-766">Selection</span></span><br>
         - <span data-ttu-id="01242-767">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-767">Settings</span></span><br>
         - <span data-ttu-id="01242-768">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-768">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-769">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="01242-769">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="01242-770">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-770">Content</span></span><br>
         - <span data-ttu-id="01242-771">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-771">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-772">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-772">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="01242-773">ActiveView</span><span class="sxs-lookup"><span data-stu-id="01242-773">ActiveView</span></span><br>
         - <span data-ttu-id="01242-774">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="01242-774">CompressedFile</span></span><br>
         - <span data-ttu-id="01242-775">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-775">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-776">File</span><span class="sxs-lookup"><span data-stu-id="01242-776">File</span></span><br>
         - <span data-ttu-id="01242-777">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-777">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-778">PdfFile</span><span class="sxs-lookup"><span data-stu-id="01242-778">PdfFile</span></span><br>
         - <span data-ttu-id="01242-779">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-779">Selection</span></span><br>
         - <span data-ttu-id="01242-780">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-780">Settings</span></span><br>
         - <span data-ttu-id="01242-781">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-781">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="01242-782">&ast; — Добавлены обновления после выпуска.</span><span class="sxs-lookup"><span data-stu-id="01242-782">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="onenote"></a><span data-ttu-id="01242-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="01242-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="01242-784">Платформа</span><span class="sxs-lookup"><span data-stu-id="01242-784">Platform</span></span></th>
    <th><span data-ttu-id="01242-785">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="01242-785">Extension points</span></span></th>
    <th><span data-ttu-id="01242-786">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="01242-786">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="01242-787">Общие API</span><span class="sxs-lookup"><span data-stu-id="01242-787">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="01242-788">Office Online</span></span></td>
    <td> - <span data-ttu-id="01242-789">Контент</span><span class="sxs-lookup"><span data-stu-id="01242-789">Content</span></span><br>
         - <span data-ttu-id="01242-790">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-790">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="01242-791">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-791">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets"><span data-ttu-id="01242-792">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-792">OneNoteApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-793">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-793">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-794">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="01242-794">DocumentEvents</span></span><br>
         - <span data-ttu-id="01242-795">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-795">HtmlCoercion</span></span><br>
         - <span data-ttu-id="01242-796">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-796">ImageCoercion</span></span><br>
         - <span data-ttu-id="01242-797">Settings</span><span class="sxs-lookup"><span data-stu-id="01242-797">Settings</span></span><br>
         - <span data-ttu-id="01242-798">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-798">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="01242-799">Project</span><span class="sxs-lookup"><span data-stu-id="01242-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="01242-800">Платформа</span><span class="sxs-lookup"><span data-stu-id="01242-800">Platform</span></span></th>
    <th><span data-ttu-id="01242-801">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="01242-801">Extension points</span></span></th>
    <th><span data-ttu-id="01242-802">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="01242-802">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="01242-803">Общие API</span><span class="sxs-lookup"><span data-stu-id="01242-803">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-804">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-804">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-805">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-805">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-806">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-806">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-807">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-807">Selection</span></span><br>
         - <span data-ttu-id="01242-808">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-808">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-809">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-809">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-810">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-810">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-811">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-811">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-812">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-812">Selection</span></span><br>
         - <span data-ttu-id="01242-813">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-813">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="01242-814">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="01242-814">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="01242-815">Область задач</span><span class="sxs-lookup"><span data-stu-id="01242-815">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="01242-816">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="01242-816">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="01242-817">Selection</span><span class="sxs-lookup"><span data-stu-id="01242-817">Selection</span></span><br>
         - <span data-ttu-id="01242-818">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="01242-818">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="01242-819">См. также</span><span class="sxs-lookup"><span data-stu-id="01242-819">See also</span></span>

- [<span data-ttu-id="01242-820">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="01242-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="01242-821">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="01242-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="01242-822">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="01242-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="01242-823">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="01242-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="01242-824">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="01242-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="01242-825">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="01242-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="01242-826">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="01242-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="01242-827">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="01242-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)