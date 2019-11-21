---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: ecb906e595c08b973b5146416a5317d59547ed39
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757487"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="6f480-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6f480-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="6f480-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="6f480-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="6f480-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="6f480-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="6f480-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="6f480-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="6f480-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="6f480-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="6f480-108">Excel</span><span class="sxs-lookup"><span data-stu-id="6f480-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6f480-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6f480-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6f480-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6f480-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="6f480-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-114">- TaskPane</span></span><br><span data-ttu-id="6f480-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-115">
        - Content</span></span><br><span data-ttu-id="6f480-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-116">
        - Custom Functions</span></span><br><span data-ttu-id="6f480-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="6f480-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6f480-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6f480-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6f480-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6f480-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6f480-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6f480-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6f480-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6f480-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6f480-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6f480-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6f480-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6f480-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="6f480-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="6f480-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-130">
        - BindingEvents</span></span><br><span data-ttu-id="6f480-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-131">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-132">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-133">
        - File</span></span><br><span data-ttu-id="6f480-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-134">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-136">
        - Selection</span></span><br><span data-ttu-id="6f480-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-137">
        - Settings</span></span><br><span data-ttu-id="6f480-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-138">
        - TableBindings</span></span><br><span data-ttu-id="6f480-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-139">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-140">
        - TextBindings</span></span><br><span data-ttu-id="6f480-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-142">Office on Windows</span></span><br><span data-ttu-id="6f480-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-144">- TaskPane</span></span><br><span data-ttu-id="6f480-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-145">
        - Content</span></span><br><span data-ttu-id="6f480-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-146">
        - Custom Functions</span></span><br><span data-ttu-id="6f480-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="6f480-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6f480-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6f480-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6f480-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6f480-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6f480-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6f480-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6f480-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6f480-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6f480-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6f480-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6f480-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6f480-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="6f480-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-161">
        - BindingEvents</span></span><br><span data-ttu-id="6f480-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-162">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-163">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-164">
        - File</span></span><br><span data-ttu-id="6f480-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-165">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-167">
        - Selection</span></span><br><span data-ttu-id="6f480-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-168">
        - Settings</span></span><br><span data-ttu-id="6f480-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-169">
        - TableBindings</span></span><br><span data-ttu-id="6f480-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-170">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-171">
        - TextBindings</span></span><br><span data-ttu-id="6f480-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-173">Office 2019 on Windows</span></span><br><span data-ttu-id="6f480-174">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6f480-175">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-175">- TaskPane</span></span><br><span data-ttu-id="6f480-176">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-176">
        - Content</span></span><br><span data-ttu-id="6f480-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6f480-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6f480-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6f480-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6f480-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6f480-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6f480-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6f480-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6f480-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-188">- BindingEvents</span></span><br><span data-ttu-id="6f480-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-189">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-190">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-191">
        - File</span></span><br><span data-ttu-id="6f480-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-192">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-194">
        - Selection</span></span><br><span data-ttu-id="6f480-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-195">
        - Settings</span></span><br><span data-ttu-id="6f480-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-196">
        - TableBindings</span></span><br><span data-ttu-id="6f480-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-197">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-198">
        - TextBindings</span></span><br><span data-ttu-id="6f480-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-200">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-200">Office 2016 on Windows</span></span><br><span data-ttu-id="6f480-201">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6f480-202">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-202">- TaskPane</span></span><br><span data-ttu-id="6f480-203">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-203">
        - Content</span></span></td>
    <td><span data-ttu-id="6f480-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6f480-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-207">- BindingEvents</span></span><br><span data-ttu-id="6f480-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-208">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-209">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-210">
        - File</span></span><br><span data-ttu-id="6f480-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-211">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-213">
        - Selection</span></span><br><span data-ttu-id="6f480-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-214">
        - Settings</span></span><br><span data-ttu-id="6f480-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-215">
        - TableBindings</span></span><br><span data-ttu-id="6f480-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-216">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-217">
        - TextBindings</span></span><br><span data-ttu-id="6f480-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-219">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-219">Office 2013 on Windows</span></span><br><span data-ttu-id="6f480-220">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6f480-221">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-221">
        - TaskPane</span></span><br><span data-ttu-id="6f480-222">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="6f480-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6f480-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6f480-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-225">
        - BindingEvents</span></span><br><span data-ttu-id="6f480-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-226">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-227">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-228">
        - File</span></span><br><span data-ttu-id="6f480-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-229">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-231">
        - Selection</span></span><br><span data-ttu-id="6f480-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-232">
        - Settings</span></span><br><span data-ttu-id="6f480-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-233">
        - TableBindings</span></span><br><span data-ttu-id="6f480-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-234">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-235">
        - TextBindings</span></span><br><span data-ttu-id="6f480-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-237">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="6f480-237">Office on iPad</span></span><br><span data-ttu-id="6f480-238">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="6f480-239">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-239">- TaskPane</span></span><br><span data-ttu-id="6f480-240">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-240">
        - Content</span></span></td>
    <td><span data-ttu-id="6f480-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6f480-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6f480-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6f480-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6f480-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6f480-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6f480-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6f480-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6f480-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6f480-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6f480-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6f480-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-253">- BindingEvents</span></span><br><span data-ttu-id="6f480-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-254">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-255">
        - File</span></span><br><span data-ttu-id="6f480-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-256">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-258">
        - Selection</span></span><br><span data-ttu-id="6f480-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-259">
        - Settings</span></span><br><span data-ttu-id="6f480-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-260">
        - TableBindings</span></span><br><span data-ttu-id="6f480-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-261">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-262">
        - TextBindings</span></span><br><span data-ttu-id="6f480-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-264">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-264">Office on Mac</span></span><br><span data-ttu-id="6f480-265">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="6f480-266">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-266">- TaskPane</span></span><br><span data-ttu-id="6f480-267">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-267">
        - Content</span></span><br><span data-ttu-id="6f480-268">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-268">
        - Custom Functions</span></span><br><span data-ttu-id="6f480-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6f480-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6f480-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6f480-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6f480-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6f480-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6f480-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6f480-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6f480-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="6f480-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="6f480-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="6f480-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="6f480-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="6f480-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-283">- BindingEvents</span></span><br><span data-ttu-id="6f480-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-284">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-285">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-286">
        - File</span></span><br><span data-ttu-id="6f480-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-287">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-289">
        - PdfFile</span></span><br><span data-ttu-id="6f480-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-290">
        - Selection</span></span><br><span data-ttu-id="6f480-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-291">
        - Settings</span></span><br><span data-ttu-id="6f480-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-292">
        - TableBindings</span></span><br><span data-ttu-id="6f480-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-293">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-294">
        - TextBindings</span></span><br><span data-ttu-id="6f480-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-296">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-296">Office 2019 on Mac</span></span><br><span data-ttu-id="6f480-297">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6f480-298">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-298">- TaskPane</span></span><br><span data-ttu-id="6f480-299">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-299">
        - Content</span></span><br><span data-ttu-id="6f480-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6f480-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6f480-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6f480-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6f480-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6f480-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6f480-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6f480-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6f480-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-311">- BindingEvents</span></span><br><span data-ttu-id="6f480-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-312">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-313">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-314">
        - File</span></span><br><span data-ttu-id="6f480-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-315">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-317">
        - PdfFile</span></span><br><span data-ttu-id="6f480-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-318">
        - Selection</span></span><br><span data-ttu-id="6f480-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-319">
        - Settings</span></span><br><span data-ttu-id="6f480-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-320">
        - TableBindings</span></span><br><span data-ttu-id="6f480-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-321">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-322">
        - TextBindings</span></span><br><span data-ttu-id="6f480-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-324">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-324">Office 2016 on Mac</span></span><br><span data-ttu-id="6f480-325">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="6f480-326">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-326">- TaskPane</span></span><br><span data-ttu-id="6f480-327">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-327">
        - Content</span></span></td>
    <td><span data-ttu-id="6f480-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6f480-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6f480-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="6f480-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-331">- BindingEvents</span></span><br><span data-ttu-id="6f480-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-332">
        - CompressedFile</span></span><br><span data-ttu-id="6f480-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-333">
        - DocumentEvents</span></span><br><span data-ttu-id="6f480-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="6f480-334">
        - File</span></span><br><span data-ttu-id="6f480-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-335">
        - MatrixBindings</span></span><br><span data-ttu-id="6f480-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="6f480-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-337">
        - PdfFile</span></span><br><span data-ttu-id="6f480-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-338">
        - Selection</span></span><br><span data-ttu-id="6f480-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-339">
        - Settings</span></span><br><span data-ttu-id="6f480-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-340">
        - TableBindings</span></span><br><span data-ttu-id="6f480-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-341">
        - TableCoercion</span></span><br><span data-ttu-id="6f480-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-342">
        - TextBindings</span></span><br><span data-ttu-id="6f480-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="6f480-344">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6f480-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="6f480-345">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6f480-346">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6f480-347">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6f480-348">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6f480-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-350">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-350">Office on the web</span></span></td>
    <td><span data-ttu-id="6f480-351">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="6f480-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-353">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-353">Office on Windows</span></span><br><span data-ttu-id="6f480-354">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="6f480-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="6f480-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-357">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-357">Office for Mac</span></span><br><span data-ttu-id="6f480-358">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="6f480-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="6f480-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="6f480-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="6f480-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="6f480-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6f480-362">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-362">Platform</span></span></th>
    <th><span data-ttu-id="6f480-363">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-363">Extension points</span></span></th>
    <th><span data-ttu-id="6f480-364">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="6f480-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-366">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-366">Office on the web</span></span><br><span data-ttu-id="6f480-367">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="6f480-367">(modern)</span></span></td>
    <td> <span data-ttu-id="6f480-368">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-368">- Mail Read</span></span><br><span data-ttu-id="6f480-369">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-369">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6f480-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6f480-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="6f480-379">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-380">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-380">Office on the web</span></span><br><span data-ttu-id="6f480-381">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="6f480-381">(classic)</span></span></td>
    <td> <span data-ttu-id="6f480-382">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-382">- Mail Read</span></span><br><span data-ttu-id="6f480-383">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-383">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6f480-391">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-392">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-392">Office on Windows</span></span><br><span data-ttu-id="6f480-393">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-394">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-394">- Mail Read</span></span><br><span data-ttu-id="6f480-395">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-395">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6f480-397">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="6f480-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6f480-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6f480-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6f480-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="6f480-406">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-407">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-407">Office 2019 on Windows</span></span><br><span data-ttu-id="6f480-408">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-409">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-409">- Mail Read</span></span><br><span data-ttu-id="6f480-410">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-410">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6f480-412">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="6f480-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6f480-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6f480-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="6f480-420">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-421">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-421">Office 2016 on Windows</span></span><br><span data-ttu-id="6f480-422">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-423">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-423">- Mail Read</span></span><br><span data-ttu-id="6f480-424">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-424">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6f480-426">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="6f480-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6f480-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="6f480-431">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-432">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-432">Office 2013 on Windows</span></span><br><span data-ttu-id="6f480-433">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-434">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-434">- Mail Read</span></span><br><span data-ttu-id="6f480-435">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="6f480-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="6f480-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="6f480-440">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-441">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="6f480-441">Office on iOS</span></span><br><span data-ttu-id="6f480-442">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-443">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-443">- Mail Read</span></span><br><span data-ttu-id="6f480-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6f480-450">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-451">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-451">Office on Mac</span></span><br><span data-ttu-id="6f480-452">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-453">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-453">- Mail Read</span></span><br><span data-ttu-id="6f480-454">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-454">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6f480-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6f480-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="6f480-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6f480-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="6f480-464">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-465">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-465">Office 2019 on Mac</span></span><br><span data-ttu-id="6f480-466">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-467">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-467">- Mail Read</span></span><br><span data-ttu-id="6f480-468">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-468">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6f480-476">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-477">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-477">Office 2016 on Mac</span></span><br><span data-ttu-id="6f480-478">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-479">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-479">- Mail Read</span></span><br><span data-ttu-id="6f480-480">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="6f480-480">
      - Mail Compose</span></span><br><span data-ttu-id="6f480-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6f480-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6f480-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6f480-488">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-489">Office для Android</span><span class="sxs-lookup"><span data-stu-id="6f480-489">Office on Android</span></span><br><span data-ttu-id="6f480-490">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-491">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="6f480-491">- Mail Read</span></span><br><span data-ttu-id="6f480-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6f480-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6f480-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6f480-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6f480-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6f480-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6f480-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6f480-498">Недоступно</span><span class="sxs-lookup"><span data-stu-id="6f480-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="6f480-499">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6f480-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6f480-500">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="6f480-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="6f480-501">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="6f480-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="6f480-502">Word</span><span class="sxs-lookup"><span data-stu-id="6f480-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6f480-503">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-503">Platform</span></span></th>
    <th><span data-ttu-id="6f480-504">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-504">Extension points</span></span></th>
    <th><span data-ttu-id="6f480-505">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="6f480-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-507">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="6f480-508">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-508">- TaskPane</span></span><br><span data-ttu-id="6f480-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6f480-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6f480-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6f480-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-516">- BindingEvents</span></span><br><span data-ttu-id="6f480-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-518">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-519">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-519">
         - File</span></span><br><span data-ttu-id="6f480-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-521">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-524">
         - PdfFile</span></span><br><span data-ttu-id="6f480-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-525">
         - Selection</span></span><br><span data-ttu-id="6f480-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-526">
         - Settings</span></span><br><span data-ttu-id="6f480-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-527">
         - TableBindings</span></span><br><span data-ttu-id="6f480-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-528">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-529">
         - TextBindings</span></span><br><span data-ttu-id="6f480-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-530">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-532">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-532">Office on Windows</span></span><br><span data-ttu-id="6f480-533">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-534">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-534">- TaskPane</span></span><br><span data-ttu-id="6f480-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6f480-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6f480-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6f480-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-542">- BindingEvents</span></span><br><span data-ttu-id="6f480-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-543">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-545">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-546">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-546">
         - File</span></span><br><span data-ttu-id="6f480-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-548">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-551">
         - PdfFile</span></span><br><span data-ttu-id="6f480-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-552">
         - Selection</span></span><br><span data-ttu-id="6f480-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-553">
         - Settings</span></span><br><span data-ttu-id="6f480-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-554">
         - TableBindings</span></span><br><span data-ttu-id="6f480-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-555">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-556">
         - TextBindings</span></span><br><span data-ttu-id="6f480-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-557">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-559">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-559">Office 2019 on Windows</span></span><br><span data-ttu-id="6f480-560">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-561">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-561">- TaskPane</span></span><br><span data-ttu-id="6f480-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6f480-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6f480-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-568">- BindingEvents</span></span><br><span data-ttu-id="6f480-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-569">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-571">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-572">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-572">
         - File</span></span><br><span data-ttu-id="6f480-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-574">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-577">
         - PdfFile</span></span><br><span data-ttu-id="6f480-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-578">
         - Selection</span></span><br><span data-ttu-id="6f480-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-579">
         - Settings</span></span><br><span data-ttu-id="6f480-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-580">
         - TableBindings</span></span><br><span data-ttu-id="6f480-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-581">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-582">
         - TextBindings</span></span><br><span data-ttu-id="6f480-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-583">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-585">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-585">Office 2016 on Windows</span></span><br><span data-ttu-id="6f480-586">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-587">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6f480-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-591">- BindingEvents</span></span><br><span data-ttu-id="6f480-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-592">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-594">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-595">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-595">
         - File</span></span><br><span data-ttu-id="6f480-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-597">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-600">
         - PdfFile</span></span><br><span data-ttu-id="6f480-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-601">
         - Selection</span></span><br><span data-ttu-id="6f480-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-602">
         - Settings</span></span><br><span data-ttu-id="6f480-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-603">
         - TableBindings</span></span><br><span data-ttu-id="6f480-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-604">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-605">
         - TextBindings</span></span><br><span data-ttu-id="6f480-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-606">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-608">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-608">Office 2013 on Windows</span></span><br><span data-ttu-id="6f480-609">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-610">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6f480-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6f480-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-613">- BindingEvents</span></span><br><span data-ttu-id="6f480-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-614">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-616">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-617">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-617">
         - File</span></span><br><span data-ttu-id="6f480-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-619">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-622">
         - PdfFile</span></span><br><span data-ttu-id="6f480-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-623">
         - Selection</span></span><br><span data-ttu-id="6f480-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-624">
         - Settings</span></span><br><span data-ttu-id="6f480-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-625">
         - TableBindings</span></span><br><span data-ttu-id="6f480-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-626">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-627">
         - TextBindings</span></span><br><span data-ttu-id="6f480-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-628">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-630">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="6f480-630">Office on iPad</span></span><br><span data-ttu-id="6f480-631">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-632">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6f480-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6f480-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="6f480-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-638">- BindingEvents</span></span><br><span data-ttu-id="6f480-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-639">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-641">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-642">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-642">
         - File</span></span><br><span data-ttu-id="6f480-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-644">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-647">
         - PdfFile</span></span><br><span data-ttu-id="6f480-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-648">
         - Selection</span></span><br><span data-ttu-id="6f480-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-649">
         - Settings</span></span><br><span data-ttu-id="6f480-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-650">
         - TableBindings</span></span><br><span data-ttu-id="6f480-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-651">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-652">
         - TextBindings</span></span><br><span data-ttu-id="6f480-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-653">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-655">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-655">Office on Mac</span></span><br><span data-ttu-id="6f480-656">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-657">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-657">- TaskPane</span></span><br><span data-ttu-id="6f480-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6f480-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6f480-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="6f480-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-665">- BindingEvents</span></span><br><span data-ttu-id="6f480-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-666">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-668">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-669">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-669">
         - File</span></span><br><span data-ttu-id="6f480-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-671">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-674">
         - PdfFile</span></span><br><span data-ttu-id="6f480-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-675">
         - Selection</span></span><br><span data-ttu-id="6f480-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-676">
         - Settings</span></span><br><span data-ttu-id="6f480-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-677">
         - TableBindings</span></span><br><span data-ttu-id="6f480-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-678">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-679">
         - TextBindings</span></span><br><span data-ttu-id="6f480-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-680">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-682">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-682">Office 2019 on Mac</span></span><br><span data-ttu-id="6f480-683">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-684">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-684">- TaskPane</span></span><br><span data-ttu-id="6f480-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="6f480-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6f480-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="6f480-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="6f480-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-691">- BindingEvents</span></span><br><span data-ttu-id="6f480-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-692">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-694">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-695">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-695">
         - File</span></span><br><span data-ttu-id="6f480-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-697">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-700">
         - PdfFile</span></span><br><span data-ttu-id="6f480-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-701">
         - Selection</span></span><br><span data-ttu-id="6f480-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-702">
         - Settings</span></span><br><span data-ttu-id="6f480-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-703">
         - TableBindings</span></span><br><span data-ttu-id="6f480-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-704">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-705">
         - TextBindings</span></span><br><span data-ttu-id="6f480-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-706">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-708">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-708">Office 2016 on Mac</span></span><br><span data-ttu-id="6f480-709">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-710">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="6f480-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6f480-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="6f480-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-714">- BindingEvents</span></span><br><span data-ttu-id="6f480-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-715">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6f480-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="6f480-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-717">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-718">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="6f480-718">
         - File</span></span><br><span data-ttu-id="6f480-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-720">
         - MatrixBindings</span></span><br><span data-ttu-id="6f480-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="6f480-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6f480-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-723">
         - PdfFile</span></span><br><span data-ttu-id="6f480-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-724">
         - Selection</span></span><br><span data-ttu-id="6f480-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="6f480-725">
         - Settings</span></span><br><span data-ttu-id="6f480-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-726">
         - TableBindings</span></span><br><span data-ttu-id="6f480-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-727">
         - TableCoercion</span></span><br><span data-ttu-id="6f480-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6f480-728">
         - TextBindings</span></span><br><span data-ttu-id="6f480-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-729">
         - TextCoercion</span></span><br><span data-ttu-id="6f480-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6f480-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="6f480-731">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6f480-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="6f480-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6f480-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6f480-733">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-733">Platform</span></span></th>
    <th><span data-ttu-id="6f480-734">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-734">Extension points</span></span></th>
    <th><span data-ttu-id="6f480-735">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="6f480-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-737">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="6f480-738">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-738">- Content</span></span><br><span data-ttu-id="6f480-739">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-739">
         - TaskPane</span></span><br><span data-ttu-id="6f480-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6f480-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6f480-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-745">- ActiveView</span></span><br><span data-ttu-id="6f480-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-746">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-747">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-748">
         - File</span></span><br><span data-ttu-id="6f480-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-749">
         - PdfFile</span></span><br><span data-ttu-id="6f480-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-750">
         - Selection</span></span><br><span data-ttu-id="6f480-751">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-751">
         - Settings</span></span><br><span data-ttu-id="6f480-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-753">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-753">Office on Windows</span></span><br><span data-ttu-id="6f480-754">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-755">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-755">- Content</span></span><br><span data-ttu-id="6f480-756">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-756">
         - TaskPane</span></span><br><span data-ttu-id="6f480-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6f480-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6f480-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-762">- ActiveView</span></span><br><span data-ttu-id="6f480-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-763">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-764">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-765">
         - File</span></span><br><span data-ttu-id="6f480-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-766">
         - PdfFile</span></span><br><span data-ttu-id="6f480-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-767">
         - Selection</span></span><br><span data-ttu-id="6f480-768">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-768">
         - Settings</span></span><br><span data-ttu-id="6f480-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-770">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-770">Office 2019 on Windows</span></span><br><span data-ttu-id="6f480-771">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-772">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-772">- Content</span></span><br><span data-ttu-id="6f480-773">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-773">
         - TaskPane</span></span><br><span data-ttu-id="6f480-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-777">- ActiveView</span></span><br><span data-ttu-id="6f480-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-778">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-779">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-780">
         - File</span></span><br><span data-ttu-id="6f480-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-781">
         - PdfFile</span></span><br><span data-ttu-id="6f480-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-782">
         - Selection</span></span><br><span data-ttu-id="6f480-783">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-783">
         - Settings</span></span><br><span data-ttu-id="6f480-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-785">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-785">Office 2016 on Windows</span></span><br><span data-ttu-id="6f480-786">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-787">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-787">- Content</span></span><br><span data-ttu-id="6f480-788">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6f480-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6f480-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-791">- ActiveView</span></span><br><span data-ttu-id="6f480-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-792">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-793">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-794">
         - File</span></span><br><span data-ttu-id="6f480-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-795">
         - PdfFile</span></span><br><span data-ttu-id="6f480-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-796">
         - Selection</span></span><br><span data-ttu-id="6f480-797">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-797">
         - Settings</span></span><br><span data-ttu-id="6f480-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-799">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-799">Office 2013 on Windows</span></span><br><span data-ttu-id="6f480-800">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-801">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-801">- Content</span></span><br><span data-ttu-id="6f480-802">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="6f480-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6f480-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6f480-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-805">- ActiveView</span></span><br><span data-ttu-id="6f480-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-806">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-807">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-808">
         - File</span></span><br><span data-ttu-id="6f480-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-809">
         - PdfFile</span></span><br><span data-ttu-id="6f480-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-810">
         - Selection</span></span><br><span data-ttu-id="6f480-811">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-811">
         - Settings</span></span><br><span data-ttu-id="6f480-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-813">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="6f480-813">Office on iPad</span></span><br><span data-ttu-id="6f480-814">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-815">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-815">- Content</span></span><br><span data-ttu-id="6f480-816">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6f480-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-820">- ActiveView</span></span><br><span data-ttu-id="6f480-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-821">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-822">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-823">
         - File</span></span><br><span data-ttu-id="6f480-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-824">
         - PdfFile</span></span><br><span data-ttu-id="6f480-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-825">
         - Selection</span></span><br><span data-ttu-id="6f480-826">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-826">
         - Settings</span></span><br><span data-ttu-id="6f480-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-828">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-828">Office on Mac</span></span><br><span data-ttu-id="6f480-829">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="6f480-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="6f480-830">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-830">- Content</span></span><br><span data-ttu-id="6f480-831">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-831">
         - TaskPane</span></span><br><span data-ttu-id="6f480-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="6f480-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="6f480-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6f480-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="6f480-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-837">- ActiveView</span></span><br><span data-ttu-id="6f480-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-838">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-839">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-840">
         - File</span></span><br><span data-ttu-id="6f480-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-841">
         - PdfFile</span></span><br><span data-ttu-id="6f480-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-842">
         - Selection</span></span><br><span data-ttu-id="6f480-843">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-843">
         - Settings</span></span><br><span data-ttu-id="6f480-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-845">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-845">Office 2019 on Mac</span></span><br><span data-ttu-id="6f480-846">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-847">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-847">- Content</span></span><br><span data-ttu-id="6f480-848">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-848">
         - TaskPane</span></span><br><span data-ttu-id="6f480-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-852">- ActiveView</span></span><br><span data-ttu-id="6f480-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-853">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-854">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-855">
         - File</span></span><br><span data-ttu-id="6f480-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-856">
         - PdfFile</span></span><br><span data-ttu-id="6f480-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-857">
         - Selection</span></span><br><span data-ttu-id="6f480-858">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-858">
         - Settings</span></span><br><span data-ttu-id="6f480-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-860">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-860">Office 2016 on Mac</span></span><br><span data-ttu-id="6f480-861">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-862">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-862">- Content</span></span><br><span data-ttu-id="6f480-863">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6f480-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="6f480-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6f480-866">- ActiveView</span></span><br><span data-ttu-id="6f480-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6f480-867">
         - CompressedFile</span></span><br><span data-ttu-id="6f480-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-868">
         - DocumentEvents</span></span><br><span data-ttu-id="6f480-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="6f480-869">
         - File</span></span><br><span data-ttu-id="6f480-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6f480-870">
         - PdfFile</span></span><br><span data-ttu-id="6f480-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-871">
         - Selection</span></span><br><span data-ttu-id="6f480-872">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-872">
         - Settings</span></span><br><span data-ttu-id="6f480-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="6f480-874">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="6f480-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="6f480-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="6f480-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6f480-876">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-876">Platform</span></span></th>
    <th><span data-ttu-id="6f480-877">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-877">Extension points</span></span></th>
    <th><span data-ttu-id="6f480-878">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="6f480-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-880">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="6f480-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="6f480-881">- Контент</span><span class="sxs-lookup"><span data-stu-id="6f480-881">- Content</span></span><br><span data-ttu-id="6f480-882">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-882">
         - TaskPane</span></span><br><span data-ttu-id="6f480-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="6f480-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6f480-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="6f480-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="6f480-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6f480-887">- DocumentEvents</span></span><br><span data-ttu-id="6f480-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="6f480-889">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="6f480-889">
         - Settings</span></span><br><span data-ttu-id="6f480-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="6f480-891">Project</span><span class="sxs-lookup"><span data-stu-id="6f480-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6f480-892">Платформа</span><span class="sxs-lookup"><span data-stu-id="6f480-892">Platform</span></span></th>
    <th><span data-ttu-id="6f480-893">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="6f480-893">Extension points</span></span></th>
    <th><span data-ttu-id="6f480-894">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="6f480-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="6f480-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="6f480-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-896">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-896">Office 2019 on Windows</span></span><br><span data-ttu-id="6f480-897">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-898">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-900">- Selection</span></span><br><span data-ttu-id="6f480-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-902">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-902">Office 2016 on Windows</span></span><br><span data-ttu-id="6f480-903">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-904">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-906">- Selection</span></span><br><span data-ttu-id="6f480-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6f480-908">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="6f480-908">Office 2013 on Windows</span></span><br><span data-ttu-id="6f480-909">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="6f480-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="6f480-910">- Область задач</span><span class="sxs-lookup"><span data-stu-id="6f480-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6f480-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6f480-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6f480-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="6f480-912">- Selection</span></span><br><span data-ttu-id="6f480-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6f480-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="6f480-914">См. также</span><span class="sxs-lookup"><span data-stu-id="6f480-914">See also</span></span>

- [<span data-ttu-id="6f480-915">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6f480-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="6f480-916">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="6f480-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="6f480-917">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="6f480-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="6f480-918">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="6f480-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="6f480-919">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="6f480-919">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="6f480-920">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="6f480-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="6f480-921">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="6f480-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="6f480-922">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="6f480-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="6f480-923">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6f480-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="6f480-924">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6f480-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="6f480-925">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="6f480-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
