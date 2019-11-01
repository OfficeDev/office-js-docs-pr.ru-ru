---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: 3621236ea86410d70d17655450e1f6d32a212823
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901950"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="09d19-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="09d19-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="09d19-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="09d19-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="09d19-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="09d19-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="09d19-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="09d19-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="09d19-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="09d19-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="09d19-108">Excel</span><span class="sxs-lookup"><span data-stu-id="09d19-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="09d19-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="09d19-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="09d19-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="09d19-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="09d19-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-114">- TaskPane</span></span><br><span data-ttu-id="09d19-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-115">
        - Content</span></span><br><span data-ttu-id="09d19-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-116">
        - Custom Functions</span></span><br><span data-ttu-id="09d19-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="09d19-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="09d19-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="09d19-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="09d19-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="09d19-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="09d19-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="09d19-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="09d19-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="09d19-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="09d19-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="09d19-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-128">
        - BindingEvents</span></span><br><span data-ttu-id="09d19-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-129">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-130">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-131">
        - File</span></span><br><span data-ttu-id="09d19-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-132">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-134">
        - Selection</span></span><br><span data-ttu-id="09d19-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-135">
        - Settings</span></span><br><span data-ttu-id="09d19-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-136">
        - TableBindings</span></span><br><span data-ttu-id="09d19-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-137">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-138">
        - TextBindings</span></span><br><span data-ttu-id="09d19-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-140">Office on Windows</span></span><br><span data-ttu-id="09d19-141">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-142">- TaskPane</span></span><br><span data-ttu-id="09d19-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-143">
        - Content</span></span><br><span data-ttu-id="09d19-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-144">
        - Custom Functions</span></span><br><span data-ttu-id="09d19-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="09d19-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="09d19-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="09d19-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="09d19-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="09d19-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="09d19-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="09d19-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="09d19-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="09d19-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="09d19-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="09d19-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="09d19-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-158">
        - BindingEvents</span></span><br><span data-ttu-id="09d19-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-159">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-160">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-161">
        - File</span></span><br><span data-ttu-id="09d19-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-162">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-164">
        - Selection</span></span><br><span data-ttu-id="09d19-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-165">
        - Settings</span></span><br><span data-ttu-id="09d19-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-166">
        - TableBindings</span></span><br><span data-ttu-id="09d19-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-167">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-168">
        - TextBindings</span></span><br><span data-ttu-id="09d19-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-170">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-170">Office 2019 on Windows</span></span><br><span data-ttu-id="09d19-171">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="09d19-172">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-172">- TaskPane</span></span><br><span data-ttu-id="09d19-173">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-173">
        - Content</span></span><br><span data-ttu-id="09d19-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="09d19-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="09d19-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="09d19-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="09d19-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="09d19-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="09d19-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="09d19-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="09d19-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-185">- BindingEvents</span></span><br><span data-ttu-id="09d19-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-186">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-187">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-188">
        - File</span></span><br><span data-ttu-id="09d19-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-189">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-191">
        - Selection</span></span><br><span data-ttu-id="09d19-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-192">
        - Settings</span></span><br><span data-ttu-id="09d19-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-193">
        - TableBindings</span></span><br><span data-ttu-id="09d19-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-194">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-195">
        - TextBindings</span></span><br><span data-ttu-id="09d19-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-197">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-197">Office 2016 on Windows</span></span><br><span data-ttu-id="09d19-198">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="09d19-199">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-199">- TaskPane</span></span><br><span data-ttu-id="09d19-200">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-200">
        - Content</span></span></td>
    <td><span data-ttu-id="09d19-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="09d19-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-204">- BindingEvents</span></span><br><span data-ttu-id="09d19-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-205">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-206">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-207">
        - File</span></span><br><span data-ttu-id="09d19-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-208">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-210">
        - Selection</span></span><br><span data-ttu-id="09d19-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-211">
        - Settings</span></span><br><span data-ttu-id="09d19-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-212">
        - TableBindings</span></span><br><span data-ttu-id="09d19-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-213">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-214">
        - TextBindings</span></span><br><span data-ttu-id="09d19-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-216">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-216">Office 2013 on Windows</span></span><br><span data-ttu-id="09d19-217">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="09d19-218">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-218">
        - TaskPane</span></span><br><span data-ttu-id="09d19-219">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="09d19-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="09d19-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="09d19-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-222">
        - BindingEvents</span></span><br><span data-ttu-id="09d19-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-223">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-224">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-225">
        - File</span></span><br><span data-ttu-id="09d19-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-226">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-228">
        - Selection</span></span><br><span data-ttu-id="09d19-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-229">
        - Settings</span></span><br><span data-ttu-id="09d19-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-230">
        - TableBindings</span></span><br><span data-ttu-id="09d19-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-231">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-232">
        - TextBindings</span></span><br><span data-ttu-id="09d19-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-234">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="09d19-234">Office on iPad</span></span><br><span data-ttu-id="09d19-235">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="09d19-236">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-236">- TaskPane</span></span><br><span data-ttu-id="09d19-237">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-237">
        - Content</span></span></td>
    <td><span data-ttu-id="09d19-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="09d19-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="09d19-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="09d19-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="09d19-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="09d19-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="09d19-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="09d19-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="09d19-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="09d19-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-249">- BindingEvents</span></span><br><span data-ttu-id="09d19-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-250">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-251">
        - File</span></span><br><span data-ttu-id="09d19-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-252">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-254">
        - Selection</span></span><br><span data-ttu-id="09d19-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-255">
        - Settings</span></span><br><span data-ttu-id="09d19-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-256">
        - TableBindings</span></span><br><span data-ttu-id="09d19-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-257">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-258">
        - TextBindings</span></span><br><span data-ttu-id="09d19-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-260">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-260">Office on Mac</span></span><br><span data-ttu-id="09d19-261">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="09d19-262">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-262">- TaskPane</span></span><br><span data-ttu-id="09d19-263">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-263">
        - Content</span></span><br><span data-ttu-id="09d19-264">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-264">
        - Custom Functions</span></span><br><span data-ttu-id="09d19-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="09d19-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="09d19-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="09d19-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="09d19-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="09d19-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="09d19-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="09d19-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="09d19-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="09d19-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="09d19-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="09d19-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-278">- BindingEvents</span></span><br><span data-ttu-id="09d19-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-279">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-280">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-281">
        - File</span></span><br><span data-ttu-id="09d19-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-282">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-284">
        - PdfFile</span></span><br><span data-ttu-id="09d19-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-285">
        - Selection</span></span><br><span data-ttu-id="09d19-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-286">
        - Settings</span></span><br><span data-ttu-id="09d19-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-287">
        - TableBindings</span></span><br><span data-ttu-id="09d19-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-288">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-289">
        - TextBindings</span></span><br><span data-ttu-id="09d19-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-291">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-291">Office 2019 on Mac</span></span><br><span data-ttu-id="09d19-292">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="09d19-293">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-293">- TaskPane</span></span><br><span data-ttu-id="09d19-294">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-294">
        - Content</span></span><br><span data-ttu-id="09d19-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="09d19-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="09d19-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="09d19-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="09d19-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="09d19-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="09d19-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="09d19-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="09d19-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-306">- BindingEvents</span></span><br><span data-ttu-id="09d19-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-307">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-308">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-309">
        - File</span></span><br><span data-ttu-id="09d19-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-310">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-312">
        - PdfFile</span></span><br><span data-ttu-id="09d19-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-313">
        - Selection</span></span><br><span data-ttu-id="09d19-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-314">
        - Settings</span></span><br><span data-ttu-id="09d19-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-315">
        - TableBindings</span></span><br><span data-ttu-id="09d19-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-316">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-317">
        - TextBindings</span></span><br><span data-ttu-id="09d19-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-319">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-319">Office 2016 on Mac</span></span><br><span data-ttu-id="09d19-320">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="09d19-321">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-321">- TaskPane</span></span><br><span data-ttu-id="09d19-322">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-322">
        - Content</span></span></td>
    <td><span data-ttu-id="09d19-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="09d19-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="09d19-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="09d19-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-326">- BindingEvents</span></span><br><span data-ttu-id="09d19-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-327">
        - CompressedFile</span></span><br><span data-ttu-id="09d19-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-328">
        - DocumentEvents</span></span><br><span data-ttu-id="09d19-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="09d19-329">
        - File</span></span><br><span data-ttu-id="09d19-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-330">
        - MatrixBindings</span></span><br><span data-ttu-id="09d19-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="09d19-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-332">
        - PdfFile</span></span><br><span data-ttu-id="09d19-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-333">
        - Selection</span></span><br><span data-ttu-id="09d19-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-334">
        - Settings</span></span><br><span data-ttu-id="09d19-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-335">
        - TableBindings</span></span><br><span data-ttu-id="09d19-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-336">
        - TableCoercion</span></span><br><span data-ttu-id="09d19-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-337">
        - TextBindings</span></span><br><span data-ttu-id="09d19-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="09d19-339">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="09d19-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="09d19-340">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="09d19-341">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="09d19-342">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="09d19-343">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="09d19-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-345">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-345">Office on the web</span></span></td>
    <td><span data-ttu-id="09d19-346">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="09d19-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-348">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-348">Office on Windows</span></span><br><span data-ttu-id="09d19-349">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="09d19-350">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="09d19-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-352">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-352">Office for Mac</span></span><br><span data-ttu-id="09d19-353">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="09d19-354">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="09d19-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="09d19-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="09d19-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="09d19-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="09d19-357">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-357">Platform</span></span></th>
    <th><span data-ttu-id="09d19-358">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-358">Extension points</span></span></th>
    <th><span data-ttu-id="09d19-359">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="09d19-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-361">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-361">Office on the web</span></span><br><span data-ttu-id="09d19-362">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="09d19-362">(modern)</span></span></td>
    <td> <span data-ttu-id="09d19-363">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-363">- Mail Read</span></span><br><span data-ttu-id="09d19-364">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-364">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="09d19-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="09d19-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="09d19-374">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-375">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-375">Office on the web</span></span><br><span data-ttu-id="09d19-376">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="09d19-376">(classic)</span></span></td>
    <td> <span data-ttu-id="09d19-377">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-377">- Mail Read</span></span><br><span data-ttu-id="09d19-378">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-378">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="09d19-386">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-387">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-387">Office on Windows</span></span><br><span data-ttu-id="09d19-388">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-389">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-389">- Mail Read</span></span><br><span data-ttu-id="09d19-390">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-390">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="09d19-392">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="09d19-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="09d19-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="09d19-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="09d19-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="09d19-401">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-401">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-402">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-402">Office 2019 on Windows</span></span><br><span data-ttu-id="09d19-403">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-403">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-404">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-404">- Mail Read</span></span><br><span data-ttu-id="09d19-405">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-405">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="09d19-407">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="09d19-407">
      - Modules</span></span></td>
    <td> <span data-ttu-id="09d19-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="09d19-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="09d19-415">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-416">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-416">Office 2016 on Windows</span></span><br><span data-ttu-id="09d19-417">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-418">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-418">- Mail Read</span></span><br><span data-ttu-id="09d19-419">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-419">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="09d19-421">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="09d19-421">
      - Modules</span></span></td>
    <td> <span data-ttu-id="09d19-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="09d19-426">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-427">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-427">Office 2013 on Windows</span></span><br><span data-ttu-id="09d19-428">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-428">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-429">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-429">- Mail Read</span></span><br><span data-ttu-id="09d19-430">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-430">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="09d19-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="09d19-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="09d19-435">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-436">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="09d19-436">Office on iOS</span></span><br><span data-ttu-id="09d19-437">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-437">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-438">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-438">- Mail Read</span></span><br><span data-ttu-id="09d19-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="09d19-445">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-446">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-446">Office on Mac</span></span><br><span data-ttu-id="09d19-447">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-447">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-448">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-448">- Mail Read</span></span><br><span data-ttu-id="09d19-449">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-449">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="09d19-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="09d19-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="09d19-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="09d19-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="09d19-459">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-460">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-460">Office 2019 on Mac</span></span><br><span data-ttu-id="09d19-461">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-462">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-462">- Mail Read</span></span><br><span data-ttu-id="09d19-463">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-463">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="09d19-471">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-472">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-472">Office 2016 on Mac</span></span><br><span data-ttu-id="09d19-473">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-474">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-474">- Mail Read</span></span><br><span data-ttu-id="09d19-475">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="09d19-475">
      - Mail Compose</span></span><br><span data-ttu-id="09d19-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="09d19-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="09d19-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="09d19-483">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-484">Office для Android</span><span class="sxs-lookup"><span data-stu-id="09d19-484">Office on Android</span></span><br><span data-ttu-id="09d19-485">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-486">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="09d19-486">- Mail Read</span></span><br><span data-ttu-id="09d19-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="09d19-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="09d19-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="09d19-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="09d19-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="09d19-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="09d19-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="09d19-493">Недоступно</span><span class="sxs-lookup"><span data-stu-id="09d19-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="09d19-494">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="09d19-494">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="09d19-495">Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="09d19-495">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="09d19-496">Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="09d19-496">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="09d19-497">Word</span><span class="sxs-lookup"><span data-stu-id="09d19-497">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="09d19-498">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-498">Platform</span></span></th>
    <th><span data-ttu-id="09d19-499">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-499">Extension points</span></span></th>
    <th><span data-ttu-id="09d19-500">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-500">API requirement sets</span></span></th>
    <th><span data-ttu-id="09d19-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-502">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-502">Office on the web</span></span></td>
    <td> <span data-ttu-id="09d19-503">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-503">- TaskPane</span></span><br><span data-ttu-id="09d19-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="09d19-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="09d19-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="09d19-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-511">- BindingEvents</span></span><br><span data-ttu-id="09d19-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-513">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-514">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-514">
         - File</span></span><br><span data-ttu-id="09d19-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-516">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-519">
         - PdfFile</span></span><br><span data-ttu-id="09d19-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-520">
         - Selection</span></span><br><span data-ttu-id="09d19-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-521">
         - Settings</span></span><br><span data-ttu-id="09d19-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-522">
         - TableBindings</span></span><br><span data-ttu-id="09d19-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-523">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-524">
         - TextBindings</span></span><br><span data-ttu-id="09d19-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-525">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-526">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-527">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-527">Office on Windows</span></span><br><span data-ttu-id="09d19-528">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-528">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-529">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-529">- TaskPane</span></span><br><span data-ttu-id="09d19-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="09d19-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="09d19-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="09d19-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-537">- BindingEvents</span></span><br><span data-ttu-id="09d19-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-538">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-540">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-541">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-541">
         - File</span></span><br><span data-ttu-id="09d19-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-543">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-546">
         - PdfFile</span></span><br><span data-ttu-id="09d19-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-547">
         - Selection</span></span><br><span data-ttu-id="09d19-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-548">
         - Settings</span></span><br><span data-ttu-id="09d19-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-549">
         - TableBindings</span></span><br><span data-ttu-id="09d19-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-550">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-551">
         - TextBindings</span></span><br><span data-ttu-id="09d19-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-552">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-553">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-554">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-554">Office 2019 on Windows</span></span><br><span data-ttu-id="09d19-555">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-555">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-556">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-556">- TaskPane</span></span><br><span data-ttu-id="09d19-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="09d19-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="09d19-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-563">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-563">- BindingEvents</span></span><br><span data-ttu-id="09d19-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-564">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-565">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-565">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-566">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-567">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-567">
         - File</span></span><br><span data-ttu-id="09d19-568">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-568">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-569">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-569">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-570">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-570">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-571">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-571">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-572">
         - PdfFile</span></span><br><span data-ttu-id="09d19-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-573">
         - Selection</span></span><br><span data-ttu-id="09d19-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-574">
         - Settings</span></span><br><span data-ttu-id="09d19-575">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-575">
         - TableBindings</span></span><br><span data-ttu-id="09d19-576">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-576">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-577">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-577">
         - TextBindings</span></span><br><span data-ttu-id="09d19-578">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-578">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-579">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-579">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-580">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-580">Office 2016 on Windows</span></span><br><span data-ttu-id="09d19-581">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-581">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-582">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-582">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="09d19-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-586">- BindingEvents</span></span><br><span data-ttu-id="09d19-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-587">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-589">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-590">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-590">
         - File</span></span><br><span data-ttu-id="09d19-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-592">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-595">
         - PdfFile</span></span><br><span data-ttu-id="09d19-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-596">
         - Selection</span></span><br><span data-ttu-id="09d19-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-597">
         - Settings</span></span><br><span data-ttu-id="09d19-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-598">
         - TableBindings</span></span><br><span data-ttu-id="09d19-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-599">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-600">
         - TextBindings</span></span><br><span data-ttu-id="09d19-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-601">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-603">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-603">Office 2013 on Windows</span></span><br><span data-ttu-id="09d19-604">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-605">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="09d19-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="09d19-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-608">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-608">- BindingEvents</span></span><br><span data-ttu-id="09d19-609">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-609">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-610">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-610">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-611">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-611">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-612">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-612">
         - File</span></span><br><span data-ttu-id="09d19-613">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-613">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-614">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-614">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-615">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-615">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-616">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-616">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-617">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-617">
         - PdfFile</span></span><br><span data-ttu-id="09d19-618">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-618">
         - Selection</span></span><br><span data-ttu-id="09d19-619">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-619">
         - Settings</span></span><br><span data-ttu-id="09d19-620">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-620">
         - TableBindings</span></span><br><span data-ttu-id="09d19-621">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-621">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-622">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-622">
         - TextBindings</span></span><br><span data-ttu-id="09d19-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-623">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-624">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-624">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-625">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="09d19-625">Office on iPad</span></span><br><span data-ttu-id="09d19-626">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-626">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-627">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-627">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="09d19-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="09d19-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="09d19-633">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-633">- BindingEvents</span></span><br><span data-ttu-id="09d19-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-634">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-635">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-635">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-636">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-636">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-637">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-637">
         - File</span></span><br><span data-ttu-id="09d19-638">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-638">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-639">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-639">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-640">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-640">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-641">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-641">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-642">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-642">
         - PdfFile</span></span><br><span data-ttu-id="09d19-643">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-643">
         - Selection</span></span><br><span data-ttu-id="09d19-644">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-644">
         - Settings</span></span><br><span data-ttu-id="09d19-645">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-645">
         - TableBindings</span></span><br><span data-ttu-id="09d19-646">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-646">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-647">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-647">
         - TextBindings</span></span><br><span data-ttu-id="09d19-648">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-648">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-649">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-649">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-650">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-650">Office on Mac</span></span><br><span data-ttu-id="09d19-651">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-651">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-652">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-652">- TaskPane</span></span><br><span data-ttu-id="09d19-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="09d19-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="09d19-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="09d19-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-660">- BindingEvents</span></span><br><span data-ttu-id="09d19-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-661">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-663">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-664">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-664">
         - File</span></span><br><span data-ttu-id="09d19-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-666">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-669">
         - PdfFile</span></span><br><span data-ttu-id="09d19-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-670">
         - Selection</span></span><br><span data-ttu-id="09d19-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-671">
         - Settings</span></span><br><span data-ttu-id="09d19-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-672">
         - TableBindings</span></span><br><span data-ttu-id="09d19-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-673">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-674">
         - TextBindings</span></span><br><span data-ttu-id="09d19-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-675">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-677">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-677">Office 2019 on Mac</span></span><br><span data-ttu-id="09d19-678">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-678">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-679">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-679">- TaskPane</span></span><br><span data-ttu-id="09d19-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="09d19-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="09d19-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="09d19-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="09d19-686">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-686">- BindingEvents</span></span><br><span data-ttu-id="09d19-687">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-687">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-688">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-688">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-689">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-689">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-690">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-690">
         - File</span></span><br><span data-ttu-id="09d19-691">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-691">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-692">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-692">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-693">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-693">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-694">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-694">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-695">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-695">
         - PdfFile</span></span><br><span data-ttu-id="09d19-696">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-696">
         - Selection</span></span><br><span data-ttu-id="09d19-697">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-697">
         - Settings</span></span><br><span data-ttu-id="09d19-698">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-698">
         - TableBindings</span></span><br><span data-ttu-id="09d19-699">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-699">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-700">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-700">
         - TextBindings</span></span><br><span data-ttu-id="09d19-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-701">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-702">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-702">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-703">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-703">Office 2016 on Mac</span></span><br><span data-ttu-id="09d19-704">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-704">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-705">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-705">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="09d19-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="09d19-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="09d19-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-709">- BindingEvents</span></span><br><span data-ttu-id="09d19-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-710">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="09d19-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="09d19-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-712">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-713">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="09d19-713">
         - File</span></span><br><span data-ttu-id="09d19-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-715">
         - MatrixBindings</span></span><br><span data-ttu-id="09d19-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="09d19-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="09d19-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-718">
         - PdfFile</span></span><br><span data-ttu-id="09d19-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-719">
         - Selection</span></span><br><span data-ttu-id="09d19-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="09d19-720">
         - Settings</span></span><br><span data-ttu-id="09d19-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-721">
         - TableBindings</span></span><br><span data-ttu-id="09d19-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-722">
         - TableCoercion</span></span><br><span data-ttu-id="09d19-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="09d19-723">
         - TextBindings</span></span><br><span data-ttu-id="09d19-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-724">
         - TextCoercion</span></span><br><span data-ttu-id="09d19-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="09d19-725">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="09d19-726">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="09d19-726">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="09d19-727">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="09d19-727">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="09d19-728">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-728">Platform</span></span></th>
    <th><span data-ttu-id="09d19-729">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-729">Extension points</span></span></th>
    <th><span data-ttu-id="09d19-730">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-730">API requirement sets</span></span></th>
    <th><span data-ttu-id="09d19-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-732">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-732">Office on the web</span></span></td>
    <td> <span data-ttu-id="09d19-733">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-733">- Content</span></span><br><span data-ttu-id="09d19-734">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-734">
         - TaskPane</span></span><br><span data-ttu-id="09d19-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="09d19-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="09d19-740">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-740">- ActiveView</span></span><br><span data-ttu-id="09d19-741">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-741">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-742">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-742">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-743">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-743">
         - File</span></span><br><span data-ttu-id="09d19-744">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-744">
         - PdfFile</span></span><br><span data-ttu-id="09d19-745">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-745">
         - Selection</span></span><br><span data-ttu-id="09d19-746">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-746">
         - Settings</span></span><br><span data-ttu-id="09d19-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-747">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-748">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-748">Office on Windows</span></span><br><span data-ttu-id="09d19-749">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-749">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-750">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-750">- Content</span></span><br><span data-ttu-id="09d19-751">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-751">
         - TaskPane</span></span><br><span data-ttu-id="09d19-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="09d19-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="09d19-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-757">- ActiveView</span></span><br><span data-ttu-id="09d19-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-758">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-759">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-760">
         - File</span></span><br><span data-ttu-id="09d19-761">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-761">
         - PdfFile</span></span><br><span data-ttu-id="09d19-762">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-762">
         - Selection</span></span><br><span data-ttu-id="09d19-763">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-763">
         - Settings</span></span><br><span data-ttu-id="09d19-764">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-764">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-765">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-765">Office 2019 on Windows</span></span><br><span data-ttu-id="09d19-766">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-766">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-767">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-767">- Content</span></span><br><span data-ttu-id="09d19-768">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-768">
         - TaskPane</span></span><br><span data-ttu-id="09d19-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-772">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-772">- ActiveView</span></span><br><span data-ttu-id="09d19-773">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-773">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-774">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-774">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-775">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-775">
         - File</span></span><br><span data-ttu-id="09d19-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-776">
         - PdfFile</span></span><br><span data-ttu-id="09d19-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-777">
         - Selection</span></span><br><span data-ttu-id="09d19-778">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-778">
         - Settings</span></span><br><span data-ttu-id="09d19-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-780">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-780">Office 2016 on Windows</span></span><br><span data-ttu-id="09d19-781">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-782">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-782">- Content</span></span><br><span data-ttu-id="09d19-783">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-783">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="09d19-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="09d19-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-786">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-786">- ActiveView</span></span><br><span data-ttu-id="09d19-787">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-787">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-788">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-788">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-789">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-789">
         - File</span></span><br><span data-ttu-id="09d19-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-790">
         - PdfFile</span></span><br><span data-ttu-id="09d19-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-791">
         - Selection</span></span><br><span data-ttu-id="09d19-792">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-792">
         - Settings</span></span><br><span data-ttu-id="09d19-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-794">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-794">Office 2013 on Windows</span></span><br><span data-ttu-id="09d19-795">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-795">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-796">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-796">- Content</span></span><br><span data-ttu-id="09d19-797">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-797">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="09d19-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="09d19-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="09d19-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-800">- ActiveView</span></span><br><span data-ttu-id="09d19-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-801">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-802">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-803">
         - File</span></span><br><span data-ttu-id="09d19-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-804">
         - PdfFile</span></span><br><span data-ttu-id="09d19-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-805">
         - Selection</span></span><br><span data-ttu-id="09d19-806">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-806">
         - Settings</span></span><br><span data-ttu-id="09d19-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-808">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="09d19-808">Office on iPad</span></span><br><span data-ttu-id="09d19-809">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-810">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-810">- Content</span></span><br><span data-ttu-id="09d19-811">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="09d19-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-815">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-815">- ActiveView</span></span><br><span data-ttu-id="09d19-816">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-816">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-817">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-817">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-818">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-818">
         - File</span></span><br><span data-ttu-id="09d19-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-819">
         - PdfFile</span></span><br><span data-ttu-id="09d19-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-820">
         - Selection</span></span><br><span data-ttu-id="09d19-821">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-821">
         - Settings</span></span><br><span data-ttu-id="09d19-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-823">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-823">Office on Mac</span></span><br><span data-ttu-id="09d19-824">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="09d19-824">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="09d19-825">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-825">- Content</span></span><br><span data-ttu-id="09d19-826">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-826">
         - TaskPane</span></span><br><span data-ttu-id="09d19-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="09d19-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="09d19-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="09d19-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="09d19-832">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-832">- ActiveView</span></span><br><span data-ttu-id="09d19-833">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-833">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-834">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-834">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-835">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-835">
         - File</span></span><br><span data-ttu-id="09d19-836">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-836">
         - PdfFile</span></span><br><span data-ttu-id="09d19-837">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-837">
         - Selection</span></span><br><span data-ttu-id="09d19-838">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-838">
         - Settings</span></span><br><span data-ttu-id="09d19-839">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-839">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-840">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-840">Office 2019 on Mac</span></span><br><span data-ttu-id="09d19-841">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-841">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-842">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-842">- Content</span></span><br><span data-ttu-id="09d19-843">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-843">
         - TaskPane</span></span><br><span data-ttu-id="09d19-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-847">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-847">- ActiveView</span></span><br><span data-ttu-id="09d19-848">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-848">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-849">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-849">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-850">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-850">
         - File</span></span><br><span data-ttu-id="09d19-851">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-851">
         - PdfFile</span></span><br><span data-ttu-id="09d19-852">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-852">
         - Selection</span></span><br><span data-ttu-id="09d19-853">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-853">
         - Settings</span></span><br><span data-ttu-id="09d19-854">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-854">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-855">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-855">Office 2016 on Mac</span></span><br><span data-ttu-id="09d19-856">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-856">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-857">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-857">- Content</span></span><br><span data-ttu-id="09d19-858">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-858">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="09d19-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="09d19-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-861">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="09d19-861">- ActiveView</span></span><br><span data-ttu-id="09d19-862">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="09d19-862">
         - CompressedFile</span></span><br><span data-ttu-id="09d19-863">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-863">
         - DocumentEvents</span></span><br><span data-ttu-id="09d19-864">
         - File</span><span class="sxs-lookup"><span data-stu-id="09d19-864">
         - File</span></span><br><span data-ttu-id="09d19-865">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="09d19-865">
         - PdfFile</span></span><br><span data-ttu-id="09d19-866">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-866">
         - Selection</span></span><br><span data-ttu-id="09d19-867">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-867">
         - Settings</span></span><br><span data-ttu-id="09d19-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="09d19-869">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="09d19-869">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="09d19-870">OneNote</span><span class="sxs-lookup"><span data-stu-id="09d19-870">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="09d19-871">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-871">Platform</span></span></th>
    <th><span data-ttu-id="09d19-872">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-872">Extension points</span></span></th>
    <th><span data-ttu-id="09d19-873">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-873">API requirement sets</span></span></th>
    <th><span data-ttu-id="09d19-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-875">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="09d19-875">Office on the web</span></span></td>
    <td> <span data-ttu-id="09d19-876">- Контент</span><span class="sxs-lookup"><span data-stu-id="09d19-876">- Content</span></span><br><span data-ttu-id="09d19-877">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-877">
         - TaskPane</span></span><br><span data-ttu-id="09d19-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="09d19-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="09d19-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="09d19-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="09d19-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-882">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="09d19-882">- DocumentEvents</span></span><br><span data-ttu-id="09d19-883">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-883">
         - HtmlCoercion</span></span><br><span data-ttu-id="09d19-884">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="09d19-884">
         - Settings</span></span><br><span data-ttu-id="09d19-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-885">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="09d19-886">Project</span><span class="sxs-lookup"><span data-stu-id="09d19-886">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="09d19-887">Платформа</span><span class="sxs-lookup"><span data-stu-id="09d19-887">Platform</span></span></th>
    <th><span data-ttu-id="09d19-888">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="09d19-888">Extension points</span></span></th>
    <th><span data-ttu-id="09d19-889">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="09d19-889">API requirement sets</span></span></th>
    <th><span data-ttu-id="09d19-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="09d19-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-891">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-891">Office 2019 on Windows</span></span><br><span data-ttu-id="09d19-892">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-893">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-895">- Selection</span></span><br><span data-ttu-id="09d19-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-897">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-897">Office 2016 on Windows</span></span><br><span data-ttu-id="09d19-898">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-899">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-901">- Selection</span></span><br><span data-ttu-id="09d19-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-902">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="09d19-903">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="09d19-903">Office 2013 on Windows</span></span><br><span data-ttu-id="09d19-904">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="09d19-904">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="09d19-905">- Область задач</span><span class="sxs-lookup"><span data-stu-id="09d19-905">- TaskPane</span></span></td>
    <td> <span data-ttu-id="09d19-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="09d19-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="09d19-907">- Selection</span><span class="sxs-lookup"><span data-stu-id="09d19-907">- Selection</span></span><br><span data-ttu-id="09d19-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="09d19-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="09d19-909">См. также</span><span class="sxs-lookup"><span data-stu-id="09d19-909">See also</span></span>

- [<span data-ttu-id="09d19-910">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="09d19-910">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="09d19-911">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="09d19-911">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="09d19-912">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="09d19-912">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="09d19-913">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="09d19-913">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="09d19-914">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="09d19-914">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="09d19-915">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="09d19-915">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="09d19-916">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="09d19-916">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="09d19-917">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="09d19-917">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="09d19-918">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="09d19-918">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="09d19-919">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="09d19-919">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="09d19-920">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="09d19-920">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
