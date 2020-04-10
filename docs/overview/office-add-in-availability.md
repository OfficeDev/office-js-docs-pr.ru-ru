---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 04/07/2020
localization_priority: Priority
ms.openlocfilehash: 823fd53e71c71f4a845f9a7b5c6177ad3f14745f
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185619"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b1aa1-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b1aa1-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b1aa1-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="b1aa1-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="b1aa1-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="b1aa1-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b1aa1-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="b1aa1-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b1aa1-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="b1aa1-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b1aa1-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b1aa1-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b1aa1-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b1aa1-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b1aa1-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b1aa1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1aa1-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-114">- TaskPane</span></span><br><span data-ttu-id="b1aa1-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-115">
        - Content</span></span><br><span data-ttu-id="b1aa1-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="b1aa1-116">
        - Custom Functions</span></span><br><span data-ttu-id="b1aa1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="b1aa1-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b1aa1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1aa1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1aa1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1aa1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1aa1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1aa1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1aa1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1aa1-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="b1aa1-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-130">
        - BindingEvents</span></span><br><span data-ttu-id="b1aa1-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-131">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-132">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-133">
        - File</span></span><br><span data-ttu-id="b1aa1-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-134">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-136">
        - Selection</span></span><br><span data-ttu-id="b1aa1-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-137">
        - Settings</span></span><br><span data-ttu-id="b1aa1-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-138">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-139">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-140">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-142">Office on Windows</span></span><br><span data-ttu-id="b1aa1-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-144">- TaskPane</span></span><br><span data-ttu-id="b1aa1-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-145">
        - Content</span></span><br><span data-ttu-id="b1aa1-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="b1aa1-146">
        - Custom Functions</span></span><br><span data-ttu-id="b1aa1-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="b1aa1-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b1aa1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1aa1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1aa1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1aa1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1aa1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1aa1-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1aa1-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1aa1-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b1aa1-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-161">
        - BindingEvents</span></span><br><span data-ttu-id="b1aa1-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-162">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-163">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-164">
        - File</span></span><br><span data-ttu-id="b1aa1-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-165">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-167">
        - Selection</span></span><br><span data-ttu-id="b1aa1-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-168">
        - Settings</span></span><br><span data-ttu-id="b1aa1-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-169">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-170">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-171">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-173">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-173">Office 2019 on Windows</span></span><br><span data-ttu-id="b1aa1-174">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1aa1-175">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-175">- TaskPane</span></span><br><span data-ttu-id="b1aa1-176">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-176">
        - Content</span></span><br><span data-ttu-id="b1aa1-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b1aa1-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1aa1-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1aa1-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1aa1-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1aa1-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1aa1-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-188">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-189">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-190">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-191">
        - File</span></span><br><span data-ttu-id="b1aa1-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-192">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-194">
        - Selection</span></span><br><span data-ttu-id="b1aa1-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-195">
        - Settings</span></span><br><span data-ttu-id="b1aa1-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-196">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-197">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-198">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-200">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-200">Office 2016 on Windows</span></span><br><span data-ttu-id="b1aa1-201">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1aa1-202">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-202">- TaskPane</span></span><br><span data-ttu-id="b1aa1-203">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-203">
        - Content</span></span></td>
    <td><span data-ttu-id="b1aa1-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1aa1-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-207">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-208">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-209">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-210">
        - File</span></span><br><span data-ttu-id="b1aa1-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-211">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-213">
        - Selection</span></span><br><span data-ttu-id="b1aa1-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-214">
        - Settings</span></span><br><span data-ttu-id="b1aa1-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-215">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-216">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-217">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-219">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-219">Office 2013 on Windows</span></span><br><span data-ttu-id="b1aa1-220">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1aa1-221">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-221">
        - TaskPane</span></span><br><span data-ttu-id="b1aa1-222">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b1aa1-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1aa1-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-225">
        - BindingEvents</span></span><br><span data-ttu-id="b1aa1-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-226">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-227">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-228">
        - File</span></span><br><span data-ttu-id="b1aa1-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-229">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-231">
        - Selection</span></span><br><span data-ttu-id="b1aa1-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-232">
        - Settings</span></span><br><span data-ttu-id="b1aa1-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-233">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-234">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-235">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-237">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="b1aa1-237">Office on iPad</span></span><br><span data-ttu-id="b1aa1-238">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b1aa1-239">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-239">- TaskPane</span></span><br><span data-ttu-id="b1aa1-240">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-240">
        - Content</span></span></td>
    <td><span data-ttu-id="b1aa1-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1aa1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1aa1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1aa1-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1aa1-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1aa1-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1aa1-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1aa1-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-253">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-254">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-255">
        - File</span></span><br><span data-ttu-id="b1aa1-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-256">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-258">
        - Selection</span></span><br><span data-ttu-id="b1aa1-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-259">
        - Settings</span></span><br><span data-ttu-id="b1aa1-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-260">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-261">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-262">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-264">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-264">Office on Mac</span></span><br><span data-ttu-id="b1aa1-265">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b1aa1-266">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-266">- TaskPane</span></span><br><span data-ttu-id="b1aa1-267">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-267">
        - Content</span></span><br><span data-ttu-id="b1aa1-268">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="b1aa1-268">
        - Custom Functions</span></span><br><span data-ttu-id="b1aa1-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b1aa1-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1aa1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1aa1-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1aa1-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1aa1-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1aa1-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b1aa1-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b1aa1-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b1aa1-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-283">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-284">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-285">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-286">
        - File</span></span><br><span data-ttu-id="b1aa1-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-287">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-289">
        - PdfFile</span></span><br><span data-ttu-id="b1aa1-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-290">
        - Selection</span></span><br><span data-ttu-id="b1aa1-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-291">
        - Settings</span></span><br><span data-ttu-id="b1aa1-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-292">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-293">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-294">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-296">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-296">Office 2019 on Mac</span></span><br><span data-ttu-id="b1aa1-297">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1aa1-298">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-298">- TaskPane</span></span><br><span data-ttu-id="b1aa1-299">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-299">
        - Content</span></span><br><span data-ttu-id="b1aa1-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b1aa1-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b1aa1-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b1aa1-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b1aa1-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b1aa1-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b1aa1-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-311">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-312">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-313">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-314">
        - File</span></span><br><span data-ttu-id="b1aa1-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-315">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-317">
        - PdfFile</span></span><br><span data-ttu-id="b1aa1-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-318">
        - Selection</span></span><br><span data-ttu-id="b1aa1-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-319">
        - Settings</span></span><br><span data-ttu-id="b1aa1-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-320">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-321">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-322">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-324">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-324">Office 2016 on Mac</span></span><br><span data-ttu-id="b1aa1-325">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b1aa1-326">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-326">- TaskPane</span></span><br><span data-ttu-id="b1aa1-327">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-327">
        - Content</span></span></td>
    <td><span data-ttu-id="b1aa1-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1aa1-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b1aa1-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-331">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-332">
        - CompressedFile</span></span><br><span data-ttu-id="b1aa1-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-333">
        - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-334">
        - File</span></span><br><span data-ttu-id="b1aa1-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-335">
        - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-337">
        - PdfFile</span></span><br><span data-ttu-id="b1aa1-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-338">
        - Selection</span></span><br><span data-ttu-id="b1aa1-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-339">
        - Settings</span></span><br><span data-ttu-id="b1aa1-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-340">
        - TableBindings</span></span><br><span data-ttu-id="b1aa1-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-341">
        - TableCoercion</span></span><br><span data-ttu-id="b1aa1-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-342">
        - TextBindings</span></span><br><span data-ttu-id="b1aa1-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b1aa1-344">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="b1aa1-345">Пользовательские функции (только Excel)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b1aa1-346">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b1aa1-347">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b1aa1-348">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b1aa1-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-350">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-350">Office on the web</span></span></td>
    <td><span data-ttu-id="b1aa1-351">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="b1aa1-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b1aa1-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-353">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-353">Office on Windows</span></span><br><span data-ttu-id="b1aa1-354">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b1aa1-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="b1aa1-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b1aa1-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-357">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-357">Office for Mac</span></span><br><span data-ttu-id="b1aa1-358">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b1aa1-359">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="b1aa1-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b1aa1-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b1aa1-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="b1aa1-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1aa1-362">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-362">Platform</span></span></th>
    <th><span data-ttu-id="b1aa1-363">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-363">Extension points</span></span></th>
    <th><span data-ttu-id="b1aa1-364">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1aa1-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-366">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-366">Office on the web</span></span><br><span data-ttu-id="b1aa1-367">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-367">(modern)</span></span></td>
    <td> <span data-ttu-id="b1aa1-368">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-368">- Message Read</span></span><br><span data-ttu-id="b1aa1-369">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-369">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-370">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-370">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-371">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-371">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1aa1-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b1aa1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b1aa1-381">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-382">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-382">Office on the web</span></span><br><span data-ttu-id="b1aa1-383">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-383">(classic)</span></span></td>
    <td> <span data-ttu-id="b1aa1-384">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-384">- Message Read</span></span><br><span data-ttu-id="b1aa1-385">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-385">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-386">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-386">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-387">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-387">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b1aa1-395">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-396">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-396">Office on Windows</span></span><br><span data-ttu-id="b1aa1-397">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-398">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-398">- Message Read</span></span><br><span data-ttu-id="b1aa1-399">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-399">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-400">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-400">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-401">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-401">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b1aa1-403">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="b1aa1-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b1aa1-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1aa1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b1aa1-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b1aa1-412">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-413">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-413">Office 2019 on Windows</span></span><br><span data-ttu-id="b1aa1-414">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-415">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-415">- Message Read</span></span><br><span data-ttu-id="b1aa1-416">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-416">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-417">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-417">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-418">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-418">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b1aa1-420">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="b1aa1-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b1aa1-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1aa1-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b1aa1-428">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-429">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-429">Office 2016 on Windows</span></span><br><span data-ttu-id="b1aa1-430">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-431">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-431">- Message Read</span></span><br><span data-ttu-id="b1aa1-432">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-432">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-433">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-433">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-434">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-434">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b1aa1-436">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="b1aa1-436">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b1aa1-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b1aa1-441">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-442">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-442">Office 2013 on Windows</span></span><br><span data-ttu-id="b1aa1-443">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-444">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-444">- Message Read</span></span><br><span data-ttu-id="b1aa1-445">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-445">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-446">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-446">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-447">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-447">
      - Appointment Organizer (Compose)</span></span><br>
    <td> <span data-ttu-id="b1aa1-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b1aa1-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b1aa1-452">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-453">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="b1aa1-453">Office on iOS</span></span><br><span data-ttu-id="b1aa1-454">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-455">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-455">- Message Read</span></span><br><span data-ttu-id="b1aa1-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b1aa1-462">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-463">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-463">Office on Mac</span></span><br><span data-ttu-id="b1aa1-464">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-465">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-465">- Message Read</span></span><br><span data-ttu-id="b1aa1-466">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-466">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-467">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-467">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-468">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-468">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b1aa1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b1aa1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b1aa1-478">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-479">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-479">Office 2019 on Mac</span></span><br><span data-ttu-id="b1aa1-480">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-481">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-481">- Message Read</span></span><br><span data-ttu-id="b1aa1-482">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-482">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-483">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-483">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-484">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-484">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b1aa1-492">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-493">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-493">Office 2016 on Mac</span></span><br><span data-ttu-id="b1aa1-494">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-495">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-495">- Message Read</span></span><br><span data-ttu-id="b1aa1-496">
      - Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-496">
      - Message Compose</span></span><br><span data-ttu-id="b1aa1-497">
      - Участник встречи (чтение)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-497">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="b1aa1-498">
      - Организатор встречи (создание)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-498">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="b1aa1-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b1aa1-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b1aa1-506">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-507">Office для Android</span><span class="sxs-lookup"><span data-stu-id="b1aa1-507">Office on Android</span></span><br><span data-ttu-id="b1aa1-508">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-509">- Прочитанное сообщение</span><span class="sxs-lookup"><span data-stu-id="b1aa1-509">- Message Read</span></span><br><span data-ttu-id="b1aa1-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b1aa1-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b1aa1-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b1aa1-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b1aa1-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b1aa1-516">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b1aa1-516">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b1aa1-517">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-517">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1aa1-518">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="b1aa1-518">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="b1aa1-519">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="b1aa1-519">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b1aa1-520">Word</span><span class="sxs-lookup"><span data-stu-id="b1aa1-520">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1aa1-521">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-521">Platform</span></span></th>
    <th><span data-ttu-id="b1aa1-522">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-522">Extension points</span></span></th>
    <th><span data-ttu-id="b1aa1-523">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-523">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1aa1-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-525">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-525">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1aa1-526">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-526">- TaskPane</span></span><br><span data-ttu-id="b1aa1-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-534">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-536">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-537">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-537">
         - File</span></span><br><span data-ttu-id="b1aa1-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-539">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-542">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-543">
         - Selection</span></span><br><span data-ttu-id="b1aa1-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-544">
         - Settings</span></span><br><span data-ttu-id="b1aa1-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-545">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-546">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-547">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-548">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-549">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-550">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-550">Office on Windows</span></span><br><span data-ttu-id="b1aa1-551">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-551">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-552">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-552">- TaskPane</span></span><br><span data-ttu-id="b1aa1-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-560">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-561">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-563">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-564">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-564">
         - File</span></span><br><span data-ttu-id="b1aa1-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-566">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-569">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-570">
         - Selection</span></span><br><span data-ttu-id="b1aa1-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-571">
         - Settings</span></span><br><span data-ttu-id="b1aa1-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-572">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-573">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-574">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-575">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-577">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-577">Office 2019 on Windows</span></span><br><span data-ttu-id="b1aa1-578">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-579">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-579">- TaskPane</span></span><br><span data-ttu-id="b1aa1-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-586">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-587">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-589">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-590">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-590">
         - File</span></span><br><span data-ttu-id="b1aa1-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-592">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-595">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-596">
         - Selection</span></span><br><span data-ttu-id="b1aa1-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-597">
         - Settings</span></span><br><span data-ttu-id="b1aa1-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-598">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-599">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-600">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-601">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-603">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-603">Office 2016 on Windows</span></span><br><span data-ttu-id="b1aa1-604">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-605">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1aa1-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-609">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-609">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-610">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-610">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-611">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-611">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-612">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-613">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-613">
         - File</span></span><br><span data-ttu-id="b1aa1-614">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-614">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-615">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-615">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-616">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-616">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-617">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-617">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-618">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-618">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-619">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-619">
         - Selection</span></span><br><span data-ttu-id="b1aa1-620">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-620">
         - Settings</span></span><br><span data-ttu-id="b1aa1-621">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-621">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-622">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-622">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-623">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-623">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-624">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-625">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-625">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-626">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-626">Office 2013 on Windows</span></span><br><span data-ttu-id="b1aa1-627">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-627">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-628">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-628">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1aa1-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-631">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-632">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-634">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-635">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-635">
         - File</span></span><br><span data-ttu-id="b1aa1-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-637">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-640">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-641">
         - Selection</span></span><br><span data-ttu-id="b1aa1-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-642">
         - Settings</span></span><br><span data-ttu-id="b1aa1-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-643">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-644">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-645">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-646">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-647">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-648">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="b1aa1-648">Office on iPad</span></span><br><span data-ttu-id="b1aa1-649">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-650">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-650">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b1aa1-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-656">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-657">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-659">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-660">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-660">
         - File</span></span><br><span data-ttu-id="b1aa1-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-662">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-665">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-666">
         - Selection</span></span><br><span data-ttu-id="b1aa1-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-667">
         - Settings</span></span><br><span data-ttu-id="b1aa1-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-668">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-669">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-670">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-671">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-673">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-673">Office on Mac</span></span><br><span data-ttu-id="b1aa1-674">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-674">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-675">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-675">- TaskPane</span></span><br><span data-ttu-id="b1aa1-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b1aa1-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-683">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-684">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-686">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-687">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-687">
         - File</span></span><br><span data-ttu-id="b1aa1-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-689">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-692">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-693">
         - Selection</span></span><br><span data-ttu-id="b1aa1-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-694">
         - Settings</span></span><br><span data-ttu-id="b1aa1-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-695">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-696">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-697">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-698">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-700">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-700">Office 2019 on Mac</span></span><br><span data-ttu-id="b1aa1-701">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-702">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-702">- TaskPane</span></span><br><span data-ttu-id="b1aa1-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b1aa1-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b1aa1-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b1aa1-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-709">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-710">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-712">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-713">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-713">
         - File</span></span><br><span data-ttu-id="b1aa1-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-715">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-718">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-719">
         - Selection</span></span><br><span data-ttu-id="b1aa1-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-720">
         - Settings</span></span><br><span data-ttu-id="b1aa1-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-721">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-722">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-723">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-724">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-725">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-726">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-726">Office 2016 on Mac</span></span><br><span data-ttu-id="b1aa1-727">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-727">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-728">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-728">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b1aa1-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-732">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-732">- BindingEvents</span></span><br><span data-ttu-id="b1aa1-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-733">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-734">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b1aa1-734">
         - CustomXmlParts</span></span><br><span data-ttu-id="b1aa1-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-735">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-736">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="b1aa1-736">
         - File</span></span><br><span data-ttu-id="b1aa1-737">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-737">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-738">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-738">
         - MatrixBindings</span></span><br><span data-ttu-id="b1aa1-739">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-739">
         - MatrixCoercion</span></span><br><span data-ttu-id="b1aa1-740">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-740">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b1aa1-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-741">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-742">
         - Selection</span></span><br><span data-ttu-id="b1aa1-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-743">
         - Settings</span></span><br><span data-ttu-id="b1aa1-744">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-744">
         - TableBindings</span></span><br><span data-ttu-id="b1aa1-745">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-745">
         - TableCoercion</span></span><br><span data-ttu-id="b1aa1-746">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b1aa1-746">
         - TextBindings</span></span><br><span data-ttu-id="b1aa1-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-747">
         - TextCoercion</span></span><br><span data-ttu-id="b1aa1-748">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-748">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b1aa1-749">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-749">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b1aa1-750">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b1aa1-750">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1aa1-751">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-751">Platform</span></span></th>
    <th><span data-ttu-id="b1aa1-752">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-752">Extension points</span></span></th>
    <th><span data-ttu-id="b1aa1-753">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-753">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1aa1-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-755">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-755">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1aa1-756">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-756">- Content</span></span><br><span data-ttu-id="b1aa1-757">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-757">
         - TaskPane</span></span><br><span data-ttu-id="b1aa1-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-763">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-763">- ActiveView</span></span><br><span data-ttu-id="b1aa1-764">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-764">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-765">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-765">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-766">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-766">
         - File</span></span><br><span data-ttu-id="b1aa1-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-767">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-768">
         - Selection</span></span><br><span data-ttu-id="b1aa1-769">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-769">
         - Settings</span></span><br><span data-ttu-id="b1aa1-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-771">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-771">Office on Windows</span></span><br><span data-ttu-id="b1aa1-772">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-772">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-773">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-773">- Content</span></span><br><span data-ttu-id="b1aa1-774">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-774">
         - TaskPane</span></span><br><span data-ttu-id="b1aa1-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-780">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-780">- ActiveView</span></span><br><span data-ttu-id="b1aa1-781">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-781">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-782">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-782">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-783">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-783">
         - File</span></span><br><span data-ttu-id="b1aa1-784">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-784">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-785">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-785">
         - Selection</span></span><br><span data-ttu-id="b1aa1-786">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-786">
         - Settings</span></span><br><span data-ttu-id="b1aa1-787">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-787">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-788">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-788">Office 2019 on Windows</span></span><br><span data-ttu-id="b1aa1-789">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-789">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-790">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-790">- Content</span></span><br><span data-ttu-id="b1aa1-791">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-791">
         - TaskPane</span></span><br><span data-ttu-id="b1aa1-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-795">- ActiveView</span></span><br><span data-ttu-id="b1aa1-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-796">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-797">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-798">
         - File</span></span><br><span data-ttu-id="b1aa1-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-799">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-800">
         - Selection</span></span><br><span data-ttu-id="b1aa1-801">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-801">
         - Settings</span></span><br><span data-ttu-id="b1aa1-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-803">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-803">Office 2016 on Windows</span></span><br><span data-ttu-id="b1aa1-804">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-804">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-805">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-805">- Content</span></span><br><span data-ttu-id="b1aa1-806">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1aa1-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-809">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-809">- ActiveView</span></span><br><span data-ttu-id="b1aa1-810">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-810">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-811">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-811">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-812">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-812">
         - File</span></span><br><span data-ttu-id="b1aa1-813">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-813">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-814">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-814">
         - Selection</span></span><br><span data-ttu-id="b1aa1-815">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-815">
         - Settings</span></span><br><span data-ttu-id="b1aa1-816">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-816">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-817">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-817">Office 2013 on Windows</span></span><br><span data-ttu-id="b1aa1-818">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-818">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-819">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-819">- Content</span></span><br><span data-ttu-id="b1aa1-820">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-820">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b1aa1-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1aa1-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-823">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-823">- ActiveView</span></span><br><span data-ttu-id="b1aa1-824">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-824">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-825">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-825">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-826">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-826">
         - File</span></span><br><span data-ttu-id="b1aa1-827">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-827">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-828">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-828">
         - Selection</span></span><br><span data-ttu-id="b1aa1-829">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-829">
         - Settings</span></span><br><span data-ttu-id="b1aa1-830">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-830">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-831">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="b1aa1-831">Office on iPad</span></span><br><span data-ttu-id="b1aa1-832">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-832">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-833">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-833">- Content</span></span><br><span data-ttu-id="b1aa1-834">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-834">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-838">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-838">- ActiveView</span></span><br><span data-ttu-id="b1aa1-839">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-839">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-840">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-840">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-841">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-841">
         - File</span></span><br><span data-ttu-id="b1aa1-842">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-842">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-843">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-843">
         - Selection</span></span><br><span data-ttu-id="b1aa1-844">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-844">
         - Settings</span></span><br><span data-ttu-id="b1aa1-845">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-845">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-846">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-846">Office on Mac</span></span><br><span data-ttu-id="b1aa1-847">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-847">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b1aa1-848">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-848">- Content</span></span><br><span data-ttu-id="b1aa1-849">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-849">
         - TaskPane</span></span><br><span data-ttu-id="b1aa1-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b1aa1-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-855">- ActiveView</span></span><br><span data-ttu-id="b1aa1-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-856">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-857">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-858">
         - File</span></span><br><span data-ttu-id="b1aa1-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-859">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-860">
         - Selection</span></span><br><span data-ttu-id="b1aa1-861">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-861">
         - Settings</span></span><br><span data-ttu-id="b1aa1-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-862">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-863">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-863">Office 2019 on Mac</span></span><br><span data-ttu-id="b1aa1-864">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-864">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-865">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-865">- Content</span></span><br><span data-ttu-id="b1aa1-866">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-866">
         - TaskPane</span></span><br><span data-ttu-id="b1aa1-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-870">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-870">- ActiveView</span></span><br><span data-ttu-id="b1aa1-871">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-871">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-872">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-872">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-873">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-873">
         - File</span></span><br><span data-ttu-id="b1aa1-874">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-874">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-875">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-875">
         - Selection</span></span><br><span data-ttu-id="b1aa1-876">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-876">
         - Settings</span></span><br><span data-ttu-id="b1aa1-877">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-877">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-878">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-878">Office 2016 on Mac</span></span><br><span data-ttu-id="b1aa1-879">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-879">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-880">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-880">- Content</span></span><br><span data-ttu-id="b1aa1-881">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-881">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b1aa1-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-884">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b1aa1-884">- ActiveView</span></span><br><span data-ttu-id="b1aa1-885">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-885">
         - CompressedFile</span></span><br><span data-ttu-id="b1aa1-886">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-886">
         - DocumentEvents</span></span><br><span data-ttu-id="b1aa1-887">
         - File</span><span class="sxs-lookup"><span data-stu-id="b1aa1-887">
         - File</span></span><br><span data-ttu-id="b1aa1-888">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b1aa1-888">
         - PdfFile</span></span><br><span data-ttu-id="b1aa1-889">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-889">
         - Selection</span></span><br><span data-ttu-id="b1aa1-890">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-890">
         - Settings</span></span><br><span data-ttu-id="b1aa1-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b1aa1-892">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="b1aa1-892">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b1aa1-893">OneNote</span><span class="sxs-lookup"><span data-stu-id="b1aa1-893">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1aa1-894">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-894">Platform</span></span></th>
    <th><span data-ttu-id="b1aa1-895">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-895">Extension points</span></span></th>
    <th><span data-ttu-id="b1aa1-896">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-896">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1aa1-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-898">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b1aa1-898">Office on the web</span></span></td>
    <td> <span data-ttu-id="b1aa1-899">- Контент</span><span class="sxs-lookup"><span data-stu-id="b1aa1-899">- Content</span></span><br><span data-ttu-id="b1aa1-900">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-900">
         - TaskPane</span></span><br><span data-ttu-id="b1aa1-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b1aa1-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-905">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b1aa1-905">- DocumentEvents</span></span><br><span data-ttu-id="b1aa1-906">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-906">
         - HtmlCoercion</span></span><br><span data-ttu-id="b1aa1-907">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="b1aa1-907">
         - Settings</span></span><br><span data-ttu-id="b1aa1-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b1aa1-909">Project</span><span class="sxs-lookup"><span data-stu-id="b1aa1-909">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b1aa1-910">Платформа</span><span class="sxs-lookup"><span data-stu-id="b1aa1-910">Platform</span></span></th>
    <th><span data-ttu-id="b1aa1-911">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b1aa1-911">Extension points</span></span></th>
    <th><span data-ttu-id="b1aa1-912">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-912">API requirement sets</span></span></th>
    <th><span data-ttu-id="b1aa1-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-914">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-914">Office 2019 on Windows</span></span><br><span data-ttu-id="b1aa1-915">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-915">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-916">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-916">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-918">- Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-918">- Selection</span></span><br><span data-ttu-id="b1aa1-919">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-919">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-920">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-920">Office 2016 on Windows</span></span><br><span data-ttu-id="b1aa1-921">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-921">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-922">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-922">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-924">- Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-924">- Selection</span></span><br><span data-ttu-id="b1aa1-925">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-925">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b1aa1-926">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="b1aa1-926">Office 2013 on Windows</span></span><br><span data-ttu-id="b1aa1-927">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-927">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b1aa1-928">- Область задач</span><span class="sxs-lookup"><span data-stu-id="b1aa1-928">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b1aa1-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b1aa1-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b1aa1-930">- Selection</span><span class="sxs-lookup"><span data-stu-id="b1aa1-930">- Selection</span></span><br><span data-ttu-id="b1aa1-931">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b1aa1-931">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b1aa1-932">См. также</span><span class="sxs-lookup"><span data-stu-id="b1aa1-932">See also</span></span>

- [<span data-ttu-id="b1aa1-933">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b1aa1-933">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b1aa1-934">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="b1aa1-934">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b1aa1-935">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-935">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="b1aa1-936">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="b1aa1-936">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="b1aa1-937">Справочная документация по API</span><span class="sxs-lookup"><span data-stu-id="b1aa1-937">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="b1aa1-938">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="b1aa1-938">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b1aa1-939">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="b1aa1-939">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b1aa1-940">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="b1aa1-940">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b1aa1-941">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-941">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b1aa1-942">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b1aa1-942">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b1aa1-943">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="b1aa1-943">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="b1aa1-944">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b1aa1-944">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)