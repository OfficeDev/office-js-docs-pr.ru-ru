---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 8c3c187d8f9b70f40a35e3773a2267dc76decbd0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611984"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c1364-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c1364-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c1364-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="c1364-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="c1364-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="c1364-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c1364-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="c1364-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c1364-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="c1364-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c1364-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c1364-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c1364-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c1364-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c1364-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c1364-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c1364-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-114">- TaskPane</span></span><br><span data-ttu-id="c1364-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-115">
        - Content</span></span><br><span data-ttu-id="c1364-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="c1364-116">
        - Custom Functions</span></span><br><span data-ttu-id="c1364-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="c1364-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c1364-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c1364-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c1364-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c1364-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c1364-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c1364-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c1364-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c1364-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c1364-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c1364-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c1364-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c1364-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c1364-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c1364-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c1364-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c1364-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-131">
        - BindingEvents</span></span><br><span data-ttu-id="c1364-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-132">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-133">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-134">
        - File</span></span><br><span data-ttu-id="c1364-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-135">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-137">
        - Selection</span></span><br><span data-ttu-id="c1364-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-138">
        - Settings</span></span><br><span data-ttu-id="c1364-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-139">
        - TableBindings</span></span><br><span data-ttu-id="c1364-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-140">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-141">
        - TextBindings</span></span><br><span data-ttu-id="c1364-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-143">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-143">Office on Windows</span></span><br><span data-ttu-id="c1364-144">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-145">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-145">- TaskPane</span></span><br><span data-ttu-id="c1364-146">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-146">
        - Content</span></span><br><span data-ttu-id="c1364-147">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="c1364-147">
        - Custom Functions</span></span><br><span data-ttu-id="c1364-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="c1364-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c1364-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c1364-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c1364-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c1364-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c1364-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c1364-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c1364-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c1364-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c1364-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c1364-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c1364-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c1364-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c1364-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c1364-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c1364-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-163">
        - BindingEvents</span></span><br><span data-ttu-id="c1364-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-164">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-165">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-166">
        - File</span></span><br><span data-ttu-id="c1364-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-167">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-169">
        - Selection</span></span><br><span data-ttu-id="c1364-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-170">
        - Settings</span></span><br><span data-ttu-id="c1364-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-171">
        - TableBindings</span></span><br><span data-ttu-id="c1364-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-172">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-173">
        - TextBindings</span></span><br><span data-ttu-id="c1364-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-175">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-175">Office 2019 on Windows</span></span><br><span data-ttu-id="c1364-176">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c1364-177">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-177">- TaskPane</span></span><br><span data-ttu-id="c1364-178">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-178">
        - Content</span></span><br><span data-ttu-id="c1364-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c1364-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c1364-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c1364-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c1364-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c1364-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c1364-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c1364-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c1364-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-190">- BindingEvents</span></span><br><span data-ttu-id="c1364-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-191">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-192">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-193">
        - File</span></span><br><span data-ttu-id="c1364-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-194">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-196">
        - Selection</span></span><br><span data-ttu-id="c1364-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-197">
        - Settings</span></span><br><span data-ttu-id="c1364-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-198">
        - TableBindings</span></span><br><span data-ttu-id="c1364-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-199">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-200">
        - TextBindings</span></span><br><span data-ttu-id="c1364-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-202">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-202">Office 2016 on Windows</span></span><br><span data-ttu-id="c1364-203">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c1364-204">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-204">- TaskPane</span></span><br><span data-ttu-id="c1364-205">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-205">
        - Content</span></span></td>
    <td><span data-ttu-id="c1364-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c1364-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-209">- BindingEvents</span></span><br><span data-ttu-id="c1364-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-210">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-211">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-212">
        - File</span></span><br><span data-ttu-id="c1364-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-213">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-215">
        - Selection</span></span><br><span data-ttu-id="c1364-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-216">
        - Settings</span></span><br><span data-ttu-id="c1364-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-217">
        - TableBindings</span></span><br><span data-ttu-id="c1364-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-218">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-219">
        - TextBindings</span></span><br><span data-ttu-id="c1364-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-221">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-221">Office 2013 on Windows</span></span><br><span data-ttu-id="c1364-222">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c1364-223">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-223">
        - TaskPane</span></span><br><span data-ttu-id="c1364-224">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c1364-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c1364-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c1364-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-227">
        - BindingEvents</span></span><br><span data-ttu-id="c1364-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-228">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-229">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-230">
        - File</span></span><br><span data-ttu-id="c1364-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-231">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-233">
        - Selection</span></span><br><span data-ttu-id="c1364-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-234">
        - Settings</span></span><br><span data-ttu-id="c1364-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-235">
        - TableBindings</span></span><br><span data-ttu-id="c1364-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-236">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-237">
        - TextBindings</span></span><br><span data-ttu-id="c1364-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-239">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="c1364-239">Office on iPad</span></span><br><span data-ttu-id="c1364-240">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c1364-241">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-241">- TaskPane</span></span><br><span data-ttu-id="c1364-242">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-242">
        - Content</span></span></td>
    <td><span data-ttu-id="c1364-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c1364-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c1364-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c1364-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c1364-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c1364-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c1364-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c1364-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c1364-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c1364-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c1364-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c1364-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c1364-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c1364-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-256">- BindingEvents</span></span><br><span data-ttu-id="c1364-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-257">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-258">
        - File</span></span><br><span data-ttu-id="c1364-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-259">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-261">
        - Selection</span></span><br><span data-ttu-id="c1364-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-262">
        - Settings</span></span><br><span data-ttu-id="c1364-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-263">
        - TableBindings</span></span><br><span data-ttu-id="c1364-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-264">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-265">
        - TextBindings</span></span><br><span data-ttu-id="c1364-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-267">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-267">Office on Mac</span></span><br><span data-ttu-id="c1364-268">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c1364-269">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-269">- TaskPane</span></span><br><span data-ttu-id="c1364-270">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-270">
        - Content</span></span><br><span data-ttu-id="c1364-271">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="c1364-271">
        - Custom Functions</span></span><br><span data-ttu-id="c1364-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c1364-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c1364-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c1364-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c1364-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c1364-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c1364-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c1364-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c1364-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c1364-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c1364-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c1364-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c1364-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c1364-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c1364-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c1364-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-287">- BindingEvents</span></span><br><span data-ttu-id="c1364-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-288">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-289">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-290">
        - File</span></span><br><span data-ttu-id="c1364-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-291">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-293">
        - PdfFile</span></span><br><span data-ttu-id="c1364-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-294">
        - Selection</span></span><br><span data-ttu-id="c1364-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-295">
        - Settings</span></span><br><span data-ttu-id="c1364-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-296">
        - TableBindings</span></span><br><span data-ttu-id="c1364-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-297">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-298">
        - TextBindings</span></span><br><span data-ttu-id="c1364-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-300">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-300">Office 2019 on Mac</span></span><br><span data-ttu-id="c1364-301">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c1364-302">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-302">- TaskPane</span></span><br><span data-ttu-id="c1364-303">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-303">
        - Content</span></span><br><span data-ttu-id="c1364-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c1364-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c1364-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c1364-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c1364-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c1364-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c1364-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c1364-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c1364-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-315">- BindingEvents</span></span><br><span data-ttu-id="c1364-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-316">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-317">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-318">
        - File</span></span><br><span data-ttu-id="c1364-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-319">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-321">
        - PdfFile</span></span><br><span data-ttu-id="c1364-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-322">
        - Selection</span></span><br><span data-ttu-id="c1364-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-323">
        - Settings</span></span><br><span data-ttu-id="c1364-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-324">
        - TableBindings</span></span><br><span data-ttu-id="c1364-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-325">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-326">
        - TextBindings</span></span><br><span data-ttu-id="c1364-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-328">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-328">Office 2016 on Mac</span></span><br><span data-ttu-id="c1364-329">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c1364-330">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-330">- TaskPane</span></span><br><span data-ttu-id="c1364-331">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-331">
        - Content</span></span></td>
    <td><span data-ttu-id="c1364-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c1364-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c1364-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c1364-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-335">- BindingEvents</span></span><br><span data-ttu-id="c1364-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-336">
        - CompressedFile</span></span><br><span data-ttu-id="c1364-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-337">
        - DocumentEvents</span></span><br><span data-ttu-id="c1364-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="c1364-338">
        - File</span></span><br><span data-ttu-id="c1364-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-339">
        - MatrixBindings</span></span><br><span data-ttu-id="c1364-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="c1364-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-341">
        - PdfFile</span></span><br><span data-ttu-id="c1364-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-342">
        - Selection</span></span><br><span data-ttu-id="c1364-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-343">
        - Settings</span></span><br><span data-ttu-id="c1364-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-344">
        - TableBindings</span></span><br><span data-ttu-id="c1364-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-345">
        - TableCoercion</span></span><br><span data-ttu-id="c1364-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-346">
        - TextBindings</span></span><br><span data-ttu-id="c1364-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c1364-348">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c1364-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="c1364-349">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="c1364-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c1364-350">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c1364-351">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c1364-352">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c1364-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-354">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-354">Office on the web</span></span></td>
    <td><span data-ttu-id="c1364-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="c1364-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c1364-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-357">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-357">Office on Windows</span></span><br><span data-ttu-id="c1364-358">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c1364-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="c1364-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c1364-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-361">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-361">Office for Mac</span></span><br><span data-ttu-id="c1364-362">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c1364-363">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="c1364-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c1364-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c1364-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="c1364-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c1364-366">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-366">Platform</span></span></th>
    <th><span data-ttu-id="c1364-367">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-367">Extension points</span></span></th>
    <th><span data-ttu-id="c1364-368">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="c1364-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-370">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-370">Office on the web</span></span><br><span data-ttu-id="c1364-371">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="c1364-371">(modern)</span></span></td>
    <td> <span data-ttu-id="c1364-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c1364-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c1364-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c1364-385">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-386">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-386">Office on the web</span></span><br><span data-ttu-id="c1364-387">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="c1364-387">(classic)</span></span></td>
    <td> <span data-ttu-id="c1364-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c1364-399">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-400">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-400">Office on Windows</span></span><br><span data-ttu-id="c1364-401">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c1364-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="c1364-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c1364-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c1364-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c1364-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c1364-416">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-417">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-417">Office 2019 on Windows</span></span><br><span data-ttu-id="c1364-418">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c1364-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="c1364-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c1364-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c1364-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c1364-432">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-433">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-433">Office 2016 on Windows</span></span><br><span data-ttu-id="c1364-434">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c1364-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Модули</a></span><span class="sxs-lookup"><span data-stu-id="c1364-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c1364-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c1364-445">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-446">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-446">Office 2013 on Windows</span></span><br><span data-ttu-id="c1364-447">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="c1364-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c1364-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c1364-456">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-457">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="c1364-457">Office on iOS</span></span><br><span data-ttu-id="c1364-458">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c1364-466">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-467">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-467">Office on Mac</span></span><br><span data-ttu-id="c1364-468">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c1364-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c1364-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c1364-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c1364-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c1364-482">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-483">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-483">Office 2019 on Mac</span></span><br><span data-ttu-id="c1364-484">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c1364-496">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-497">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-497">Office 2016 on Mac</span></span><br><span data-ttu-id="c1364-498">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Создание сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c1364-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Участник встречи (чтение)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c1364-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Организатор встречи (создание)</a></span><span class="sxs-lookup"><span data-stu-id="c1364-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c1364-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c1364-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c1364-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c1364-510">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-511">Office для Android</span><span class="sxs-lookup"><span data-stu-id="c1364-511">Office on Android</span></span><br><span data-ttu-id="c1364-512">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Чтение сообщения</a></span><span class="sxs-lookup"><span data-stu-id="c1364-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c1364-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Организатор встречи (создание): собрание по сети</a> (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c1364-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="c1364-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c1364-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c1364-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c1364-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c1364-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c1364-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c1364-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c1364-521">Недоступно</span><span class="sxs-lookup"><span data-stu-id="c1364-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c1364-522">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c1364-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c1364-523">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="c1364-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c1364-524">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="c1364-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c1364-525">Word</span><span class="sxs-lookup"><span data-stu-id="c1364-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c1364-526">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-526">Platform</span></span></th>
    <th><span data-ttu-id="c1364-527">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-527">Extension points</span></span></th>
    <th><span data-ttu-id="c1364-528">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="c1364-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-530">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="c1364-531">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-531">- TaskPane</span></span><br><span data-ttu-id="c1364-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c1364-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c1364-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c1364-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-539">- BindingEvents</span></span><br><span data-ttu-id="c1364-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-541">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-542">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-542">
         - File</span></span><br><span data-ttu-id="c1364-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-544">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-547">
         - PdfFile</span></span><br><span data-ttu-id="c1364-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-548">
         - Selection</span></span><br><span data-ttu-id="c1364-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-549">
         - Settings</span></span><br><span data-ttu-id="c1364-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-550">
         - TableBindings</span></span><br><span data-ttu-id="c1364-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-551">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-552">
         - TextBindings</span></span><br><span data-ttu-id="c1364-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-553">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-555">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-555">Office on Windows</span></span><br><span data-ttu-id="c1364-556">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-557">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-557">- TaskPane</span></span><br><span data-ttu-id="c1364-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c1364-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c1364-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c1364-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-565">- BindingEvents</span></span><br><span data-ttu-id="c1364-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-566">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-568">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-569">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-569">
         - File</span></span><br><span data-ttu-id="c1364-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-571">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-574">
         - PdfFile</span></span><br><span data-ttu-id="c1364-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-575">
         - Selection</span></span><br><span data-ttu-id="c1364-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-576">
         - Settings</span></span><br><span data-ttu-id="c1364-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-577">
         - TableBindings</span></span><br><span data-ttu-id="c1364-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-578">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-579">
         - TextBindings</span></span><br><span data-ttu-id="c1364-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-580">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-582">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-582">Office 2019 on Windows</span></span><br><span data-ttu-id="c1364-583">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-584">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-584">- TaskPane</span></span><br><span data-ttu-id="c1364-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c1364-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c1364-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-591">- BindingEvents</span></span><br><span data-ttu-id="c1364-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-592">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-594">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-595">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-595">
         - File</span></span><br><span data-ttu-id="c1364-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-600">
         - PdfFile</span></span><br><span data-ttu-id="c1364-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-601">
         - Selection</span></span><br><span data-ttu-id="c1364-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-602">
         - Settings</span></span><br><span data-ttu-id="c1364-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-603">
         - TableBindings</span></span><br><span data-ttu-id="c1364-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-604">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-605">
         - TextBindings</span></span><br><span data-ttu-id="c1364-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-606">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-608">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-608">Office 2016 on Windows</span></span><br><span data-ttu-id="c1364-609">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-610">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c1364-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-614">- BindingEvents</span></span><br><span data-ttu-id="c1364-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-615">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-617">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-618">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-618">
         - File</span></span><br><span data-ttu-id="c1364-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-620">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-623">
         - PdfFile</span></span><br><span data-ttu-id="c1364-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-624">
         - Selection</span></span><br><span data-ttu-id="c1364-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-625">
         - Settings</span></span><br><span data-ttu-id="c1364-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-626">
         - TableBindings</span></span><br><span data-ttu-id="c1364-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-627">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-628">
         - TextBindings</span></span><br><span data-ttu-id="c1364-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-629">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-631">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-631">Office 2013 on Windows</span></span><br><span data-ttu-id="c1364-632">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-633">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c1364-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c1364-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-636">- BindingEvents</span></span><br><span data-ttu-id="c1364-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-637">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-639">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-640">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-640">
         - File</span></span><br><span data-ttu-id="c1364-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-642">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-645">
         - PdfFile</span></span><br><span data-ttu-id="c1364-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-646">
         - Selection</span></span><br><span data-ttu-id="c1364-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-647">
         - Settings</span></span><br><span data-ttu-id="c1364-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-648">
         - TableBindings</span></span><br><span data-ttu-id="c1364-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-649">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-650">
         - TextBindings</span></span><br><span data-ttu-id="c1364-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-651">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-653">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="c1364-653">Office on iPad</span></span><br><span data-ttu-id="c1364-654">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-655">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c1364-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c1364-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c1364-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-661">- BindingEvents</span></span><br><span data-ttu-id="c1364-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-662">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-664">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-665">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-665">
         - File</span></span><br><span data-ttu-id="c1364-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-667">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-670">
         - PdfFile</span></span><br><span data-ttu-id="c1364-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-671">
         - Selection</span></span><br><span data-ttu-id="c1364-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-672">
         - Settings</span></span><br><span data-ttu-id="c1364-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-673">
         - TableBindings</span></span><br><span data-ttu-id="c1364-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-674">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-675">
         - TextBindings</span></span><br><span data-ttu-id="c1364-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-676">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-678">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-678">Office on Mac</span></span><br><span data-ttu-id="c1364-679">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-680">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-680">- TaskPane</span></span><br><span data-ttu-id="c1364-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c1364-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c1364-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c1364-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-688">- BindingEvents</span></span><br><span data-ttu-id="c1364-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-689">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-691">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-692">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-692">
         - File</span></span><br><span data-ttu-id="c1364-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-694">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-697">
         - PdfFile</span></span><br><span data-ttu-id="c1364-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-698">
         - Selection</span></span><br><span data-ttu-id="c1364-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-699">
         - Settings</span></span><br><span data-ttu-id="c1364-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-700">
         - TableBindings</span></span><br><span data-ttu-id="c1364-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-701">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-702">
         - TextBindings</span></span><br><span data-ttu-id="c1364-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-703">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-705">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-705">Office 2019 on Mac</span></span><br><span data-ttu-id="c1364-706">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-707">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-707">- TaskPane</span></span><br><span data-ttu-id="c1364-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c1364-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c1364-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c1364-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c1364-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-714">- BindingEvents</span></span><br><span data-ttu-id="c1364-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-715">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-717">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-718">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-718">
         - File</span></span><br><span data-ttu-id="c1364-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-720">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-723">
         - PdfFile</span></span><br><span data-ttu-id="c1364-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-724">
         - Selection</span></span><br><span data-ttu-id="c1364-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-725">
         - Settings</span></span><br><span data-ttu-id="c1364-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-726">
         - TableBindings</span></span><br><span data-ttu-id="c1364-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-727">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-728">
         - TextBindings</span></span><br><span data-ttu-id="c1364-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-729">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-731">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-731">Office 2016 on Mac</span></span><br><span data-ttu-id="c1364-732">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-733">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c1364-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c1364-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c1364-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-737">- BindingEvents</span></span><br><span data-ttu-id="c1364-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-738">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c1364-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="c1364-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-740">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-741">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="c1364-741">
         - File</span></span><br><span data-ttu-id="c1364-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-743">
         - MatrixBindings</span></span><br><span data-ttu-id="c1364-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="c1364-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c1364-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-746">
         - PdfFile</span></span><br><span data-ttu-id="c1364-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-747">
         - Selection</span></span><br><span data-ttu-id="c1364-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c1364-748">
         - Settings</span></span><br><span data-ttu-id="c1364-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-749">
         - TableBindings</span></span><br><span data-ttu-id="c1364-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-750">
         - TableCoercion</span></span><br><span data-ttu-id="c1364-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c1364-751">
         - TextBindings</span></span><br><span data-ttu-id="c1364-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-752">
         - TextCoercion</span></span><br><span data-ttu-id="c1364-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c1364-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c1364-754">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c1364-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c1364-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c1364-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c1364-756">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-756">Platform</span></span></th>
    <th><span data-ttu-id="c1364-757">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-757">Extension points</span></span></th>
    <th><span data-ttu-id="c1364-758">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="c1364-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-760">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="c1364-761">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-761">- Content</span></span><br><span data-ttu-id="c1364-762">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-762">
         - TaskPane</span></span><br><span data-ttu-id="c1364-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c1364-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c1364-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-768">- ActiveView</span></span><br><span data-ttu-id="c1364-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-769">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-770">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-771">
         - File</span></span><br><span data-ttu-id="c1364-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-772">
         - PdfFile</span></span><br><span data-ttu-id="c1364-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-773">
         - Selection</span></span><br><span data-ttu-id="c1364-774">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-774">
         - Settings</span></span><br><span data-ttu-id="c1364-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-776">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-776">Office on Windows</span></span><br><span data-ttu-id="c1364-777">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-778">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-778">- Content</span></span><br><span data-ttu-id="c1364-779">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-779">
         - TaskPane</span></span><br><span data-ttu-id="c1364-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c1364-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c1364-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-785">- ActiveView</span></span><br><span data-ttu-id="c1364-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-786">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-787">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-788">
         - File</span></span><br><span data-ttu-id="c1364-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-789">
         - PdfFile</span></span><br><span data-ttu-id="c1364-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-790">
         - Selection</span></span><br><span data-ttu-id="c1364-791">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-791">
         - Settings</span></span><br><span data-ttu-id="c1364-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-793">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-793">Office 2019 on Windows</span></span><br><span data-ttu-id="c1364-794">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-795">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-795">- Content</span></span><br><span data-ttu-id="c1364-796">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-796">
         - TaskPane</span></span><br><span data-ttu-id="c1364-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-800">- ActiveView</span></span><br><span data-ttu-id="c1364-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-801">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-802">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-803">
         - File</span></span><br><span data-ttu-id="c1364-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-804">
         - PdfFile</span></span><br><span data-ttu-id="c1364-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-805">
         - Selection</span></span><br><span data-ttu-id="c1364-806">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-806">
         - Settings</span></span><br><span data-ttu-id="c1364-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-808">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-808">Office 2016 on Windows</span></span><br><span data-ttu-id="c1364-809">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-810">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-810">- Content</span></span><br><span data-ttu-id="c1364-811">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c1364-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c1364-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-814">- ActiveView</span></span><br><span data-ttu-id="c1364-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-815">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-816">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-817">
         - File</span></span><br><span data-ttu-id="c1364-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-818">
         - PdfFile</span></span><br><span data-ttu-id="c1364-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-819">
         - Selection</span></span><br><span data-ttu-id="c1364-820">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-820">
         - Settings</span></span><br><span data-ttu-id="c1364-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-822">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-822">Office 2013 on Windows</span></span><br><span data-ttu-id="c1364-823">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-824">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-824">- Content</span></span><br><span data-ttu-id="c1364-825">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c1364-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c1364-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c1364-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-828">- ActiveView</span></span><br><span data-ttu-id="c1364-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-829">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-830">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-831">
         - File</span></span><br><span data-ttu-id="c1364-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-832">
         - PdfFile</span></span><br><span data-ttu-id="c1364-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-833">
         - Selection</span></span><br><span data-ttu-id="c1364-834">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-834">
         - Settings</span></span><br><span data-ttu-id="c1364-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-836">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="c1364-836">Office on iPad</span></span><br><span data-ttu-id="c1364-837">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-838">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-838">- Content</span></span><br><span data-ttu-id="c1364-839">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c1364-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-843">- ActiveView</span></span><br><span data-ttu-id="c1364-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-844">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-845">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-846">
         - File</span></span><br><span data-ttu-id="c1364-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-847">
         - PdfFile</span></span><br><span data-ttu-id="c1364-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-848">
         - Selection</span></span><br><span data-ttu-id="c1364-849">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-849">
         - Settings</span></span><br><span data-ttu-id="c1364-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-851">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-851">Office on Mac</span></span><br><span data-ttu-id="c1364-852">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="c1364-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c1364-853">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-853">- Content</span></span><br><span data-ttu-id="c1364-854">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-854">
         - TaskPane</span></span><br><span data-ttu-id="c1364-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c1364-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c1364-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c1364-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c1364-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-860">- ActiveView</span></span><br><span data-ttu-id="c1364-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-861">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-862">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-863">
         - File</span></span><br><span data-ttu-id="c1364-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-864">
         - PdfFile</span></span><br><span data-ttu-id="c1364-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-865">
         - Selection</span></span><br><span data-ttu-id="c1364-866">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-866">
         - Settings</span></span><br><span data-ttu-id="c1364-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-868">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-868">Office 2019 on Mac</span></span><br><span data-ttu-id="c1364-869">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-870">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-870">- Content</span></span><br><span data-ttu-id="c1364-871">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-871">
         - TaskPane</span></span><br><span data-ttu-id="c1364-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-875">- ActiveView</span></span><br><span data-ttu-id="c1364-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-876">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-877">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-878">
         - File</span></span><br><span data-ttu-id="c1364-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-879">
         - PdfFile</span></span><br><span data-ttu-id="c1364-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-880">
         - Selection</span></span><br><span data-ttu-id="c1364-881">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-881">
         - Settings</span></span><br><span data-ttu-id="c1364-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-883">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-883">Office 2016 on Mac</span></span><br><span data-ttu-id="c1364-884">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-885">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-885">- Content</span></span><br><span data-ttu-id="c1364-886">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c1364-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c1364-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c1364-889">- ActiveView</span></span><br><span data-ttu-id="c1364-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c1364-890">
         - CompressedFile</span></span><br><span data-ttu-id="c1364-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-891">
         - DocumentEvents</span></span><br><span data-ttu-id="c1364-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="c1364-892">
         - File</span></span><br><span data-ttu-id="c1364-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c1364-893">
         - PdfFile</span></span><br><span data-ttu-id="c1364-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-894">
         - Selection</span></span><br><span data-ttu-id="c1364-895">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-895">
         - Settings</span></span><br><span data-ttu-id="c1364-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c1364-897">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="c1364-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c1364-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="c1364-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c1364-899">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-899">Platform</span></span></th>
    <th><span data-ttu-id="c1364-900">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-900">Extension points</span></span></th>
    <th><span data-ttu-id="c1364-901">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="c1364-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-903">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="c1364-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="c1364-904">- Контент</span><span class="sxs-lookup"><span data-stu-id="c1364-904">- Content</span></span><br><span data-ttu-id="c1364-905">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-905">
         - TaskPane</span></span><br><span data-ttu-id="c1364-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="c1364-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c1364-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c1364-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c1364-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c1364-910">- DocumentEvents</span></span><br><span data-ttu-id="c1364-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="c1364-912">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="c1364-912">
         - Settings</span></span><br><span data-ttu-id="c1364-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c1364-914">Project</span><span class="sxs-lookup"><span data-stu-id="c1364-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c1364-915">Платформа</span><span class="sxs-lookup"><span data-stu-id="c1364-915">Platform</span></span></th>
    <th><span data-ttu-id="c1364-916">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="c1364-916">Extension points</span></span></th>
    <th><span data-ttu-id="c1364-917">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="c1364-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="c1364-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="c1364-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-919">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-919">Office 2019 on Windows</span></span><br><span data-ttu-id="c1364-920">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-921">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-923">- Selection</span></span><br><span data-ttu-id="c1364-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-925">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-925">Office 2016 on Windows</span></span><br><span data-ttu-id="c1364-926">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-927">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-929">- Selection</span></span><br><span data-ttu-id="c1364-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c1364-931">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="c1364-931">Office 2013 on Windows</span></span><br><span data-ttu-id="c1364-932">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="c1364-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c1364-933">- Область задач</span><span class="sxs-lookup"><span data-stu-id="c1364-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c1364-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c1364-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c1364-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="c1364-935">- Selection</span></span><br><span data-ttu-id="c1364-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c1364-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c1364-937">См. также</span><span class="sxs-lookup"><span data-stu-id="c1364-937">See also</span></span>

- [<span data-ttu-id="c1364-938">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c1364-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c1364-939">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="c1364-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c1364-940">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="c1364-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c1364-941">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="c1364-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c1364-942">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="c1364-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c1364-943">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="c1364-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c1364-944">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="c1364-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c1364-945">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="c1364-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c1364-946">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c1364-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c1364-947">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c1364-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c1364-948">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="c1364-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="c1364-949">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c1364-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)