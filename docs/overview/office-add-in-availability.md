---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: 510f2419d5d364a536f8c96f2057505161f03993
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804648"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="17591-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="17591-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="17591-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="17591-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="17591-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="17591-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="17591-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="17591-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="17591-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="17591-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="17591-108">Excel</span><span class="sxs-lookup"><span data-stu-id="17591-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="17591-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="17591-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="17591-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="17591-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="17591-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-114">- TaskPane</span></span><br><span data-ttu-id="17591-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-115">
        - Content</span></span><br><span data-ttu-id="17591-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-116">
        - Custom Functions</span></span><br><span data-ttu-id="17591-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="17591-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="17591-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17591-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17591-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17591-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17591-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17591-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17591-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17591-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17591-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17591-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17591-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="17591-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-130">
        - BindingEvents</span></span><br><span data-ttu-id="17591-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-131">
        - CompressedFile</span></span><br><span data-ttu-id="17591-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-132">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-133">
        - File</span></span><br><span data-ttu-id="17591-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-134">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-136">
        - Selection</span></span><br><span data-ttu-id="17591-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-137">
        - Settings</span></span><br><span data-ttu-id="17591-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-138">
        - TableBindings</span></span><br><span data-ttu-id="17591-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-139">
        - TableCoercion</span></span><br><span data-ttu-id="17591-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-140">
        - TextBindings</span></span><br><span data-ttu-id="17591-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-142">Office on Windows</span></span><br><span data-ttu-id="17591-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-144">- TaskPane</span></span><br><span data-ttu-id="17591-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-145">
        - Content</span></span><br><span data-ttu-id="17591-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-146">
        - Custom Functions</span></span><br><span data-ttu-id="17591-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="17591-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="17591-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17591-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17591-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17591-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17591-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17591-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17591-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17591-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17591-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17591-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17591-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="17591-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-160">
        - BindingEvents</span></span><br><span data-ttu-id="17591-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-161">
        - CompressedFile</span></span><br><span data-ttu-id="17591-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-162">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-163">
        - File</span></span><br><span data-ttu-id="17591-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-164">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-166">
        - Selection</span></span><br><span data-ttu-id="17591-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-167">
        - Settings</span></span><br><span data-ttu-id="17591-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-168">
        - TableBindings</span></span><br><span data-ttu-id="17591-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-169">
        - TableCoercion</span></span><br><span data-ttu-id="17591-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-170">
        - TextBindings</span></span><br><span data-ttu-id="17591-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-172">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-172">Office 2019 on Windows</span></span><br><span data-ttu-id="17591-173">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17591-174">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-174">- TaskPane</span></span><br><span data-ttu-id="17591-175">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-175">
        - Content</span></span><br><span data-ttu-id="17591-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="17591-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17591-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17591-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17591-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17591-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17591-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17591-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17591-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17591-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="17591-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-187">- BindingEvents</span></span><br><span data-ttu-id="17591-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-188">
        - CompressedFile</span></span><br><span data-ttu-id="17591-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-189">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-190">
        - File</span></span><br><span data-ttu-id="17591-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-191">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-193">
        - Selection</span></span><br><span data-ttu-id="17591-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-194">
        - Settings</span></span><br><span data-ttu-id="17591-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-195">
        - TableBindings</span></span><br><span data-ttu-id="17591-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-196">
        - TableCoercion</span></span><br><span data-ttu-id="17591-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-197">
        - TextBindings</span></span><br><span data-ttu-id="17591-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-199">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-199">Office 2016 on Windows</span></span><br><span data-ttu-id="17591-200">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17591-201">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-201">- TaskPane</span></span><br><span data-ttu-id="17591-202">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-202">
        - Content</span></span></td>
    <td><span data-ttu-id="17591-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="17591-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="17591-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-206">- BindingEvents</span></span><br><span data-ttu-id="17591-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-207">
        - CompressedFile</span></span><br><span data-ttu-id="17591-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-208">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-209">
        - File</span></span><br><span data-ttu-id="17591-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-210">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-212">
        - Selection</span></span><br><span data-ttu-id="17591-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-213">
        - Settings</span></span><br><span data-ttu-id="17591-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-214">
        - TableBindings</span></span><br><span data-ttu-id="17591-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-215">
        - TableCoercion</span></span><br><span data-ttu-id="17591-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-216">
        - TextBindings</span></span><br><span data-ttu-id="17591-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-218">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-218">Office 2013 on Windows</span></span><br><span data-ttu-id="17591-219">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17591-220">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-220">
        - TaskPane</span></span><br><span data-ttu-id="17591-221">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="17591-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17591-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="17591-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="17591-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-224">
        - BindingEvents</span></span><br><span data-ttu-id="17591-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-225">
        - CompressedFile</span></span><br><span data-ttu-id="17591-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-226">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-227">
        - File</span></span><br><span data-ttu-id="17591-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-228">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-230">
        - Selection</span></span><br><span data-ttu-id="17591-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-231">
        - Settings</span></span><br><span data-ttu-id="17591-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-232">
        - TableBindings</span></span><br><span data-ttu-id="17591-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-233">
        - TableCoercion</span></span><br><span data-ttu-id="17591-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-234">
        - TextBindings</span></span><br><span data-ttu-id="17591-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-236">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="17591-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="17591-237">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="17591-238">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-238">- TaskPane</span></span><br><span data-ttu-id="17591-239">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-239">
        - Content</span></span><br><span data-ttu-id="17591-240">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="17591-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17591-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17591-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17591-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17591-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17591-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17591-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17591-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17591-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17591-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17591-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="17591-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-252">- BindingEvents</span></span><br><span data-ttu-id="17591-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-253">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-254">
        - File</span></span><br><span data-ttu-id="17591-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-255">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-257">
        - Selection</span></span><br><span data-ttu-id="17591-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-258">
        - Settings</span></span><br><span data-ttu-id="17591-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-259">
        - TableBindings</span></span><br><span data-ttu-id="17591-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-260">
        - TableCoercion</span></span><br><span data-ttu-id="17591-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-261">
        - TextBindings</span></span><br><span data-ttu-id="17591-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-263">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-263">Office apps on Mac</span></span><br><span data-ttu-id="17591-264">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="17591-265">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-265">- TaskPane</span></span><br><span data-ttu-id="17591-266">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-266">
        - Content</span></span><br><span data-ttu-id="17591-267">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-267">
        - Custom Functions</span></span><br><span data-ttu-id="17591-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="17591-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17591-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17591-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17591-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17591-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17591-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17591-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17591-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17591-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="17591-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="17591-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="17591-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-281">- BindingEvents</span></span><br><span data-ttu-id="17591-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-282">
        - CompressedFile</span></span><br><span data-ttu-id="17591-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-283">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-284">
        - File</span></span><br><span data-ttu-id="17591-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-285">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-287">
        - PdfFile</span></span><br><span data-ttu-id="17591-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-288">
        - Selection</span></span><br><span data-ttu-id="17591-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-289">
        - Settings</span></span><br><span data-ttu-id="17591-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-290">
        - TableBindings</span></span><br><span data-ttu-id="17591-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-291">
        - TableCoercion</span></span><br><span data-ttu-id="17591-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-292">
        - TextBindings</span></span><br><span data-ttu-id="17591-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-294">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-294">Office 2019 for Mac</span></span><br><span data-ttu-id="17591-295">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17591-296">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-296">- TaskPane</span></span><br><span data-ttu-id="17591-297">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-297">
        - Content</span></span><br><span data-ttu-id="17591-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="17591-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="17591-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="17591-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="17591-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="17591-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="17591-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="17591-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="17591-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="17591-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="17591-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-309">- BindingEvents</span></span><br><span data-ttu-id="17591-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-310">
        - CompressedFile</span></span><br><span data-ttu-id="17591-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-311">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-312">
        - File</span></span><br><span data-ttu-id="17591-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-313">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-315">
        - PdfFile</span></span><br><span data-ttu-id="17591-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-316">
        - Selection</span></span><br><span data-ttu-id="17591-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-317">
        - Settings</span></span><br><span data-ttu-id="17591-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-318">
        - TableBindings</span></span><br><span data-ttu-id="17591-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-319">
        - TableCoercion</span></span><br><span data-ttu-id="17591-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-320">
        - TextBindings</span></span><br><span data-ttu-id="17591-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-322">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="17591-323">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="17591-324">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-324">- TaskPane</span></span><br><span data-ttu-id="17591-325">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="17591-325">
        - Content</span></span></td>
    <td><span data-ttu-id="17591-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="17591-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="17591-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="17591-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-329">- BindingEvents</span></span><br><span data-ttu-id="17591-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-330">
        - CompressedFile</span></span><br><span data-ttu-id="17591-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-331">
        - DocumentEvents</span></span><br><span data-ttu-id="17591-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="17591-332">
        - File</span></span><br><span data-ttu-id="17591-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-333">
        - MatrixBindings</span></span><br><span data-ttu-id="17591-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="17591-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-335">
        - PdfFile</span></span><br><span data-ttu-id="17591-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-336">
        - Selection</span></span><br><span data-ttu-id="17591-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-337">
        - Settings</span></span><br><span data-ttu-id="17591-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-338">
        - TableBindings</span></span><br><span data-ttu-id="17591-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-339">
        - TableCoercion</span></span><br><span data-ttu-id="17591-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-340">
        - TextBindings</span></span><br><span data-ttu-id="17591-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="17591-342">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="17591-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="17591-343">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="17591-344">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="17591-345">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="17591-346">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="17591-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-348">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-348">Office on the web</span></span></td>
    <td><span data-ttu-id="17591-349">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="17591-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-351">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-351">Office on Windows</span></span><br><span data-ttu-id="17591-352">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="17591-353">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="17591-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-355">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-355">Office for Mac</span></span><br><span data-ttu-id="17591-356">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="17591-357">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="17591-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="17591-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="17591-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="17591-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17591-360">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-360">Platform</span></span></th>
    <th><span data-ttu-id="17591-361">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-361">Extension points</span></span></th>
    <th><span data-ttu-id="17591-362">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="17591-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-364">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-364">Office on the web</span></span><br><span data-ttu-id="17591-365">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="17591-365">Modern</span></span></td>
    <td> <span data-ttu-id="17591-366">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-366">- Mail Read</span></span><br><span data-ttu-id="17591-367">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-367">
      - Mail Compose</span></span><br><span data-ttu-id="17591-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17591-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17591-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-377">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-377">Office on the web</span></span><br><span data-ttu-id="17591-378">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="17591-378">(classic)</span></span></td>
    <td> <span data-ttu-id="17591-379">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-379">- Mail Read</span></span><br><span data-ttu-id="17591-380">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-380">
      - Mail Compose</span></span><br><span data-ttu-id="17591-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="17591-388">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-389">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-389">Office on Windows</span></span><br><span data-ttu-id="17591-390">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-391">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-391">- Mail Read</span></span><br><span data-ttu-id="17591-392">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-392">
      - Mail Compose</span></span><br><span data-ttu-id="17591-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="17591-394">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="17591-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="17591-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17591-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17591-402">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-403">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-403">Office 2019 on Windows</span></span><br><span data-ttu-id="17591-404">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-405">- Mail Read</span></span><br><span data-ttu-id="17591-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-406">
      - Mail Compose</span></span><br><span data-ttu-id="17591-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="17591-408">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="17591-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="17591-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17591-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17591-416">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-417">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-417">Office 2016 on Windows</span></span><br><span data-ttu-id="17591-418">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-419">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-419">- Mail Read</span></span><br><span data-ttu-id="17591-420">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-420">
      - Mail Compose</span></span><br><span data-ttu-id="17591-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="17591-422">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="17591-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="17591-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="17591-427">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-428">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-428">Office 2013 on Windows</span></span><br><span data-ttu-id="17591-429">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-430">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-430">- Mail Read</span></span><br><span data-ttu-id="17591-431">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="17591-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="17591-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="17591-436">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-437">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="17591-437">Office apps on iOS</span></span><br><span data-ttu-id="17591-438">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-439">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-439">- Mail Read</span></span><br><span data-ttu-id="17591-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="17591-446">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-447">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-447">Office apps on Mac</span></span><br><span data-ttu-id="17591-448">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-449">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-449">- Mail Read</span></span><br><span data-ttu-id="17591-450">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-450">
      - Mail Compose</span></span><br><span data-ttu-id="17591-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="17591-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="17591-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="17591-459">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-460">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-460">Office 2019 for Mac</span></span><br><span data-ttu-id="17591-461">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-462">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-462">- Mail Read</span></span><br><span data-ttu-id="17591-463">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-463">
      - Mail Compose</span></span><br><span data-ttu-id="17591-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="17591-471">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-472">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="17591-473">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-474">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-474">- Mail Read</span></span><br><span data-ttu-id="17591-475">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="17591-475">
      - Mail Compose</span></span><br><span data-ttu-id="17591-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="17591-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="17591-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="17591-483">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-484">Office для Android</span><span class="sxs-lookup"><span data-stu-id="17591-484">Office apps on Android</span></span><br><span data-ttu-id="17591-485">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-486">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="17591-486">- Mail Read</span></span><br><span data-ttu-id="17591-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="17591-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="17591-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="17591-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="17591-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="17591-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="17591-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="17591-493">Недоступно</span><span class="sxs-lookup"><span data-stu-id="17591-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="17591-494">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="17591-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="17591-495">Word</span><span class="sxs-lookup"><span data-stu-id="17591-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17591-496">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-496">Platform</span></span></th>
    <th><span data-ttu-id="17591-497">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-497">Extension points</span></span></th>
    <th><span data-ttu-id="17591-498">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="17591-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-500">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="17591-501">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-501">- TaskPane</span></span><br><span data-ttu-id="17591-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="17591-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="17591-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="17591-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-509">- BindingEvents</span></span><br><span data-ttu-id="17591-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-511">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-512">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-512">
         - File</span></span><br><span data-ttu-id="17591-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-514">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-517">
         - PdfFile</span></span><br><span data-ttu-id="17591-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-518">
         - Selection</span></span><br><span data-ttu-id="17591-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-519">
         - Settings</span></span><br><span data-ttu-id="17591-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-520">
         - TableBindings</span></span><br><span data-ttu-id="17591-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-521">
         - TableCoercion</span></span><br><span data-ttu-id="17591-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-522">
         - TextBindings</span></span><br><span data-ttu-id="17591-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-523">
         - TextCoercion</span></span><br><span data-ttu-id="17591-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-525">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-525">Office on Windows</span></span><br><span data-ttu-id="17591-526">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-527">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-527">- TaskPane</span></span><br><span data-ttu-id="17591-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="17591-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="17591-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="17591-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-535">- BindingEvents</span></span><br><span data-ttu-id="17591-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-536">
         - CompressedFile</span></span><br><span data-ttu-id="17591-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-538">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-539">
         - File</span></span><br><span data-ttu-id="17591-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-541">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-544">
         - PdfFile</span></span><br><span data-ttu-id="17591-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-545">
         - Selection</span></span><br><span data-ttu-id="17591-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-546">
         - Settings</span></span><br><span data-ttu-id="17591-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-547">
         - TableBindings</span></span><br><span data-ttu-id="17591-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-548">
         - TableCoercion</span></span><br><span data-ttu-id="17591-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-549">
         - TextBindings</span></span><br><span data-ttu-id="17591-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-550">
         - TextCoercion</span></span><br><span data-ttu-id="17591-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-552">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-552">Office 2019 on Windows</span></span><br><span data-ttu-id="17591-553">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-554">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-554">- TaskPane</span></span><br><span data-ttu-id="17591-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="17591-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="17591-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-561">- BindingEvents</span></span><br><span data-ttu-id="17591-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-562">
         - CompressedFile</span></span><br><span data-ttu-id="17591-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-564">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-565">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-565">
         - File</span></span><br><span data-ttu-id="17591-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-567">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-570">
         - PdfFile</span></span><br><span data-ttu-id="17591-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-571">
         - Selection</span></span><br><span data-ttu-id="17591-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-572">
         - Settings</span></span><br><span data-ttu-id="17591-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-573">
         - TableBindings</span></span><br><span data-ttu-id="17591-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-574">
         - TableCoercion</span></span><br><span data-ttu-id="17591-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-575">
         - TextBindings</span></span><br><span data-ttu-id="17591-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-576">
         - TextCoercion</span></span><br><span data-ttu-id="17591-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-578">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-578">Office 2016 on Windows</span></span><br><span data-ttu-id="17591-579">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-580">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="17591-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-584">- BindingEvents</span></span><br><span data-ttu-id="17591-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-585">
         - CompressedFile</span></span><br><span data-ttu-id="17591-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-587">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-588">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-588">
         - File</span></span><br><span data-ttu-id="17591-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-590">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-593">
         - PdfFile</span></span><br><span data-ttu-id="17591-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-594">
         - Selection</span></span><br><span data-ttu-id="17591-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-595">
         - Settings</span></span><br><span data-ttu-id="17591-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-596">
         - TableBindings</span></span><br><span data-ttu-id="17591-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-597">
         - TableCoercion</span></span><br><span data-ttu-id="17591-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-598">
         - TextBindings</span></span><br><span data-ttu-id="17591-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-599">
         - TextCoercion</span></span><br><span data-ttu-id="17591-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-601">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-601">Office 2013 on Windows</span></span><br><span data-ttu-id="17591-602">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-603">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17591-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="17591-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-606">- BindingEvents</span></span><br><span data-ttu-id="17591-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-607">
         - CompressedFile</span></span><br><span data-ttu-id="17591-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-609">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-610">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-610">
         - File</span></span><br><span data-ttu-id="17591-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-612">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-615">
         - PdfFile</span></span><br><span data-ttu-id="17591-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-616">
         - Selection</span></span><br><span data-ttu-id="17591-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-617">
         - Settings</span></span><br><span data-ttu-id="17591-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-618">
         - TableBindings</span></span><br><span data-ttu-id="17591-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-619">
         - TableCoercion</span></span><br><span data-ttu-id="17591-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-620">
         - TextBindings</span></span><br><span data-ttu-id="17591-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-621">
         - TextCoercion</span></span><br><span data-ttu-id="17591-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-623">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="17591-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="17591-624">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-625">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="17591-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="17591-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="17591-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-631">- BindingEvents</span></span><br><span data-ttu-id="17591-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-632">
         - CompressedFile</span></span><br><span data-ttu-id="17591-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-634">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-635">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-635">
         - File</span></span><br><span data-ttu-id="17591-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-637">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-640">
         - PdfFile</span></span><br><span data-ttu-id="17591-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-641">
         - Selection</span></span><br><span data-ttu-id="17591-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-642">
         - Settings</span></span><br><span data-ttu-id="17591-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-643">
         - TableBindings</span></span><br><span data-ttu-id="17591-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-644">
         - TableCoercion</span></span><br><span data-ttu-id="17591-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-645">
         - TextBindings</span></span><br><span data-ttu-id="17591-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-646">
         - TextCoercion</span></span><br><span data-ttu-id="17591-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-648">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-648">Office apps on Mac</span></span><br><span data-ttu-id="17591-649">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-650">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-650">- TaskPane</span></span><br><span data-ttu-id="17591-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="17591-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="17591-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="17591-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-658">- BindingEvents</span></span><br><span data-ttu-id="17591-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-659">
         - CompressedFile</span></span><br><span data-ttu-id="17591-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-661">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-662">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-662">
         - File</span></span><br><span data-ttu-id="17591-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-664">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-667">
         - PdfFile</span></span><br><span data-ttu-id="17591-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-668">
         - Selection</span></span><br><span data-ttu-id="17591-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-669">
         - Settings</span></span><br><span data-ttu-id="17591-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-670">
         - TableBindings</span></span><br><span data-ttu-id="17591-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-671">
         - TableCoercion</span></span><br><span data-ttu-id="17591-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-672">
         - TextBindings</span></span><br><span data-ttu-id="17591-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-673">
         - TextCoercion</span></span><br><span data-ttu-id="17591-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-675">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-675">Office 2019 for Mac</span></span><br><span data-ttu-id="17591-676">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-677">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-677">- TaskPane</span></span><br><span data-ttu-id="17591-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="17591-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="17591-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="17591-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="17591-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-684">- BindingEvents</span></span><br><span data-ttu-id="17591-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-685">
         - CompressedFile</span></span><br><span data-ttu-id="17591-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-687">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-688">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-688">
         - File</span></span><br><span data-ttu-id="17591-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-690">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-693">
         - PdfFile</span></span><br><span data-ttu-id="17591-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-694">
         - Selection</span></span><br><span data-ttu-id="17591-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-695">
         - Settings</span></span><br><span data-ttu-id="17591-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-696">
         - TableBindings</span></span><br><span data-ttu-id="17591-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-697">
         - TableCoercion</span></span><br><span data-ttu-id="17591-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-698">
         - TextBindings</span></span><br><span data-ttu-id="17591-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-699">
         - TextCoercion</span></span><br><span data-ttu-id="17591-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-701">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="17591-702">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-703">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="17591-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="17591-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="17591-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="17591-707">- BindingEvents</span></span><br><span data-ttu-id="17591-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-708">
         - CompressedFile</span></span><br><span data-ttu-id="17591-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="17591-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="17591-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-710">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-711">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="17591-711">
         - File</span></span><br><span data-ttu-id="17591-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="17591-713">
         - MatrixBindings</span></span><br><span data-ttu-id="17591-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="17591-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="17591-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-716">
         - PdfFile</span></span><br><span data-ttu-id="17591-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-717">
         - Selection</span></span><br><span data-ttu-id="17591-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="17591-718">
         - Settings</span></span><br><span data-ttu-id="17591-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="17591-719">
         - TableBindings</span></span><br><span data-ttu-id="17591-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-720">
         - TableCoercion</span></span><br><span data-ttu-id="17591-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="17591-721">
         - TextBindings</span></span><br><span data-ttu-id="17591-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-722">
         - TextCoercion</span></span><br><span data-ttu-id="17591-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="17591-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="17591-724">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="17591-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="17591-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="17591-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17591-726">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-726">Platform</span></span></th>
    <th><span data-ttu-id="17591-727">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-727">Extension points</span></span></th>
    <th><span data-ttu-id="17591-728">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="17591-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-730">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="17591-731">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-731">- Content</span></span><br><span data-ttu-id="17591-732">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-732">
         - TaskPane</span></span><br><span data-ttu-id="17591-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="17591-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-737">- ActiveView</span></span><br><span data-ttu-id="17591-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-738">
         - CompressedFile</span></span><br><span data-ttu-id="17591-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-739">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-740">
         - File</span></span><br><span data-ttu-id="17591-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-741">
         - PdfFile</span></span><br><span data-ttu-id="17591-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-742">
         - Selection</span></span><br><span data-ttu-id="17591-743">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-743">
         - Settings</span></span><br><span data-ttu-id="17591-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-745">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-745">Office on Windows</span></span><br><span data-ttu-id="17591-746">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-747">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-747">- Content</span></span><br><span data-ttu-id="17591-748">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-748">
         - TaskPane</span></span><br><span data-ttu-id="17591-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="17591-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-753">- ActiveView</span></span><br><span data-ttu-id="17591-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-754">
         - CompressedFile</span></span><br><span data-ttu-id="17591-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-755">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-756">
         - File</span></span><br><span data-ttu-id="17591-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-757">
         - PdfFile</span></span><br><span data-ttu-id="17591-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-758">
         - Selection</span></span><br><span data-ttu-id="17591-759">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-759">
         - Settings</span></span><br><span data-ttu-id="17591-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-761">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-761">Office 2019 on Windows</span></span><br><span data-ttu-id="17591-762">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-763">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-763">- Content</span></span><br><span data-ttu-id="17591-764">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-764">
         - TaskPane</span></span><br><span data-ttu-id="17591-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-768">- ActiveView</span></span><br><span data-ttu-id="17591-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-769">
         - CompressedFile</span></span><br><span data-ttu-id="17591-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-770">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-771">
         - File</span></span><br><span data-ttu-id="17591-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-772">
         - PdfFile</span></span><br><span data-ttu-id="17591-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-773">
         - Selection</span></span><br><span data-ttu-id="17591-774">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-774">
         - Settings</span></span><br><span data-ttu-id="17591-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-776">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-776">Office 2016 on Windows</span></span><br><span data-ttu-id="17591-777">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-778">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-778">- Content</span></span><br><span data-ttu-id="17591-779">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17591-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="17591-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-782">- ActiveView</span></span><br><span data-ttu-id="17591-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-783">
         - CompressedFile</span></span><br><span data-ttu-id="17591-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-784">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-785">
         - File</span></span><br><span data-ttu-id="17591-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-786">
         - PdfFile</span></span><br><span data-ttu-id="17591-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-787">
         - Selection</span></span><br><span data-ttu-id="17591-788">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-788">
         - Settings</span></span><br><span data-ttu-id="17591-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-790">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-790">Office 2013 on Windows</span></span><br><span data-ttu-id="17591-791">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-792">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-792">- Content</span></span><br><span data-ttu-id="17591-793">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="17591-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17591-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="17591-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-796">- ActiveView</span></span><br><span data-ttu-id="17591-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-797">
         - CompressedFile</span></span><br><span data-ttu-id="17591-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-798">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-799">
         - File</span></span><br><span data-ttu-id="17591-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-800">
         - PdfFile</span></span><br><span data-ttu-id="17591-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-801">
         - Selection</span></span><br><span data-ttu-id="17591-802">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-802">
         - Settings</span></span><br><span data-ttu-id="17591-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-804">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="17591-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="17591-805">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-806">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-806">- Content</span></span><br><span data-ttu-id="17591-807">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-810">- ActiveView</span></span><br><span data-ttu-id="17591-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-811">
         - CompressedFile</span></span><br><span data-ttu-id="17591-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-812">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-813">
         - File</span></span><br><span data-ttu-id="17591-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-814">
         - PdfFile</span></span><br><span data-ttu-id="17591-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-815">
         - Selection</span></span><br><span data-ttu-id="17591-816">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-816">
         - Settings</span></span><br><span data-ttu-id="17591-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-818">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-818">Office apps on Mac</span></span><br><span data-ttu-id="17591-819">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="17591-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="17591-820">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-820">- Content</span></span><br><span data-ttu-id="17591-821">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-821">
         - TaskPane</span></span><br><span data-ttu-id="17591-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="17591-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="17591-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="17591-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-826">- ActiveView</span></span><br><span data-ttu-id="17591-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-827">
         - CompressedFile</span></span><br><span data-ttu-id="17591-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-828">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-829">
         - File</span></span><br><span data-ttu-id="17591-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-830">
         - PdfFile</span></span><br><span data-ttu-id="17591-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-831">
         - Selection</span></span><br><span data-ttu-id="17591-832">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-832">
         - Settings</span></span><br><span data-ttu-id="17591-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-834">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-834">Office 2019 for Mac</span></span><br><span data-ttu-id="17591-835">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-836">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-836">- Content</span></span><br><span data-ttu-id="17591-837">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-837">
         - TaskPane</span></span><br><span data-ttu-id="17591-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-841">- ActiveView</span></span><br><span data-ttu-id="17591-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-842">
         - CompressedFile</span></span><br><span data-ttu-id="17591-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-843">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-844">
         - File</span></span><br><span data-ttu-id="17591-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-845">
         - PdfFile</span></span><br><span data-ttu-id="17591-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-846">
         - Selection</span></span><br><span data-ttu-id="17591-847">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-847">
         - Settings</span></span><br><span data-ttu-id="17591-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-849">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-849">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="17591-850">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-851">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-851">- Content</span></span><br><span data-ttu-id="17591-852">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="17591-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="17591-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="17591-855">- ActiveView</span></span><br><span data-ttu-id="17591-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="17591-856">
         - CompressedFile</span></span><br><span data-ttu-id="17591-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-857">
         - DocumentEvents</span></span><br><span data-ttu-id="17591-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="17591-858">
         - File</span></span><br><span data-ttu-id="17591-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="17591-859">
         - PdfFile</span></span><br><span data-ttu-id="17591-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="17591-860">
         - Selection</span></span><br><span data-ttu-id="17591-861">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-861">
         - Settings</span></span><br><span data-ttu-id="17591-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="17591-863">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="17591-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="17591-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="17591-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17591-865">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-865">Platform</span></span></th>
    <th><span data-ttu-id="17591-866">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-866">Extension points</span></span></th>
    <th><span data-ttu-id="17591-867">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="17591-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-869">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="17591-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="17591-870">- Контент</span><span class="sxs-lookup"><span data-stu-id="17591-870">- Content</span></span><br><span data-ttu-id="17591-871">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-871">
         - TaskPane</span></span><br><span data-ttu-id="17591-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="17591-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="17591-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="17591-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="17591-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="17591-876">- DocumentEvents</span></span><br><span data-ttu-id="17591-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="17591-878">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="17591-878">
         - Settings</span></span><br><span data-ttu-id="17591-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="17591-880">Project</span><span class="sxs-lookup"><span data-stu-id="17591-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="17591-881">Платформа</span><span class="sxs-lookup"><span data-stu-id="17591-881">Platform</span></span></th>
    <th><span data-ttu-id="17591-882">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="17591-882">Extension points</span></span></th>
    <th><span data-ttu-id="17591-883">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="17591-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="17591-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="17591-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-885">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-885">Office 2019 on Windows</span></span><br><span data-ttu-id="17591-886">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-887">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="17591-889">- Selection</span></span><br><span data-ttu-id="17591-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-891">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-891">Office 2016 on Windows</span></span><br><span data-ttu-id="17591-892">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-893">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="17591-895">- Selection</span></span><br><span data-ttu-id="17591-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="17591-897">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="17591-897">Office 2013 on Windows</span></span><br><span data-ttu-id="17591-898">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="17591-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="17591-899">- Область задач</span><span class="sxs-lookup"><span data-stu-id="17591-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="17591-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="17591-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="17591-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="17591-901">- Selection</span></span><br><span data-ttu-id="17591-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="17591-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="17591-903">См. также</span><span class="sxs-lookup"><span data-stu-id="17591-903">See also</span></span>

- [<span data-ttu-id="17591-904">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="17591-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="17591-905">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="17591-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="17591-906">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="17591-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="17591-907">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="17591-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="17591-908">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="17591-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="17591-909">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="17591-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="17591-910">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="17591-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="17591-911">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="17591-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="17591-912">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="17591-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="17591-913">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="17591-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="17591-914">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="17591-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
