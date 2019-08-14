---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: 1e368fe21a1bcdb2a7f44c88ce8e881605fa96f2
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395654"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="50980-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="50980-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="50980-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="50980-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="50980-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="50980-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="50980-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="50980-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="50980-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="50980-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="50980-108">Excel</span><span class="sxs-lookup"><span data-stu-id="50980-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="50980-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="50980-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="50980-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="50980-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="50980-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-114">- TaskPane</span></span><br><span data-ttu-id="50980-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-115">
        - Content</span></span><br><span data-ttu-id="50980-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-116">
        - Custom Functions</span></span><br><span data-ttu-id="50980-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="50980-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="50980-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="50980-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="50980-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="50980-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="50980-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="50980-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="50980-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="50980-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="50980-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="50980-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="50980-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-128">
        - BindingEvents</span></span><br><span data-ttu-id="50980-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-129">
        - CompressedFile</span></span><br><span data-ttu-id="50980-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-130">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-131">
        - File</span></span><br><span data-ttu-id="50980-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-132">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-134">
        - Selection</span></span><br><span data-ttu-id="50980-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-135">
        - Settings</span></span><br><span data-ttu-id="50980-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-136">
        - TableBindings</span></span><br><span data-ttu-id="50980-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-137">
        - TableCoercion</span></span><br><span data-ttu-id="50980-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-138">
        - TextBindings</span></span><br><span data-ttu-id="50980-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-140">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-140">Office on Windows</span></span><br><span data-ttu-id="50980-141">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-142">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-142">- TaskPane</span></span><br><span data-ttu-id="50980-143">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-143">
        - Content</span></span><br><span data-ttu-id="50980-144">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-144">
        - Custom Functions</span></span><br><span data-ttu-id="50980-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="50980-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="50980-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="50980-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="50980-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="50980-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="50980-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="50980-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="50980-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="50980-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="50980-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="50980-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="50980-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="50980-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-158">
        - BindingEvents</span></span><br><span data-ttu-id="50980-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-159">
        - CompressedFile</span></span><br><span data-ttu-id="50980-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-160">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-161">
        - File</span></span><br><span data-ttu-id="50980-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-162">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-164">
        - Selection</span></span><br><span data-ttu-id="50980-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-165">
        - Settings</span></span><br><span data-ttu-id="50980-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-166">
        - TableBindings</span></span><br><span data-ttu-id="50980-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-167">
        - TableCoercion</span></span><br><span data-ttu-id="50980-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-168">
        - TextBindings</span></span><br><span data-ttu-id="50980-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-170">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-170">Office 2019 on Windows</span></span><br><span data-ttu-id="50980-171">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="50980-172">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-172">- TaskPane</span></span><br><span data-ttu-id="50980-173">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-173">
        - Content</span></span><br><span data-ttu-id="50980-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="50980-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="50980-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="50980-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="50980-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="50980-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="50980-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="50980-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="50980-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="50980-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-185">- BindingEvents</span></span><br><span data-ttu-id="50980-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-186">
        - CompressedFile</span></span><br><span data-ttu-id="50980-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-187">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-188">
        - File</span></span><br><span data-ttu-id="50980-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-189">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-191">
        - Selection</span></span><br><span data-ttu-id="50980-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-192">
        - Settings</span></span><br><span data-ttu-id="50980-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-193">
        - TableBindings</span></span><br><span data-ttu-id="50980-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-194">
        - TableCoercion</span></span><br><span data-ttu-id="50980-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-195">
        - TextBindings</span></span><br><span data-ttu-id="50980-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-197">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-197">Office 2016 on Windows</span></span><br><span data-ttu-id="50980-198">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="50980-199">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-199">- TaskPane</span></span><br><span data-ttu-id="50980-200">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-200">
        - Content</span></span></td>
    <td><span data-ttu-id="50980-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="50980-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-204">- BindingEvents</span></span><br><span data-ttu-id="50980-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-205">
        - CompressedFile</span></span><br><span data-ttu-id="50980-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-206">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-207">
        - File</span></span><br><span data-ttu-id="50980-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-208">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-210">
        - Selection</span></span><br><span data-ttu-id="50980-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-211">
        - Settings</span></span><br><span data-ttu-id="50980-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-212">
        - TableBindings</span></span><br><span data-ttu-id="50980-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-213">
        - TableCoercion</span></span><br><span data-ttu-id="50980-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-214">
        - TextBindings</span></span><br><span data-ttu-id="50980-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-216">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-216">Office 2013 on Windows</span></span><br><span data-ttu-id="50980-217">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="50980-218">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-218">
        - TaskPane</span></span><br><span data-ttu-id="50980-219">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="50980-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="50980-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="50980-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-222">
        - BindingEvents</span></span><br><span data-ttu-id="50980-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-223">
        - CompressedFile</span></span><br><span data-ttu-id="50980-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-224">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-225">
        - File</span></span><br><span data-ttu-id="50980-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-226">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-228">
        - Selection</span></span><br><span data-ttu-id="50980-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-229">
        - Settings</span></span><br><span data-ttu-id="50980-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-230">
        - TableBindings</span></span><br><span data-ttu-id="50980-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-231">
        - TableCoercion</span></span><br><span data-ttu-id="50980-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-232">
        - TextBindings</span></span><br><span data-ttu-id="50980-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-234">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="50980-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="50980-235">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="50980-236">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-236">- TaskPane</span></span><br><span data-ttu-id="50980-237">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-237">
        - Content</span></span><br><span data-ttu-id="50980-238">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-238">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="50980-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="50980-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="50980-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="50980-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="50980-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="50980-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="50980-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="50980-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="50980-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="50980-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="50980-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-250">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-250">- BindingEvents</span></span><br><span data-ttu-id="50980-251">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-251">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-252">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-252">
        - File</span></span><br><span data-ttu-id="50980-253">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-253">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-254">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-254">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-255">
        - Selection</span></span><br><span data-ttu-id="50980-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-256">
        - Settings</span></span><br><span data-ttu-id="50980-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-257">
        - TableBindings</span></span><br><span data-ttu-id="50980-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-258">
        - TableCoercion</span></span><br><span data-ttu-id="50980-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-259">
        - TextBindings</span></span><br><span data-ttu-id="50980-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-260">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-261">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-261">Office apps on Mac</span></span><br><span data-ttu-id="50980-262">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-262">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="50980-263">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-263">- TaskPane</span></span><br><span data-ttu-id="50980-264">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-264">
        - Content</span></span><br><span data-ttu-id="50980-265">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-265">
        - Custom Functions</span></span><br><span data-ttu-id="50980-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="50980-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="50980-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="50980-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="50980-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="50980-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="50980-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="50980-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="50980-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="50980-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="50980-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="50980-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="50980-279">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-279">- BindingEvents</span></span><br><span data-ttu-id="50980-280">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-280">
        - CompressedFile</span></span><br><span data-ttu-id="50980-281">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-281">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-282">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-282">
        - File</span></span><br><span data-ttu-id="50980-283">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-283">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-284">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-284">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-285">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-285">
        - PdfFile</span></span><br><span data-ttu-id="50980-286">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-286">
        - Selection</span></span><br><span data-ttu-id="50980-287">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-287">
        - Settings</span></span><br><span data-ttu-id="50980-288">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-288">
        - TableBindings</span></span><br><span data-ttu-id="50980-289">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-289">
        - TableCoercion</span></span><br><span data-ttu-id="50980-290">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-290">
        - TextBindings</span></span><br><span data-ttu-id="50980-291">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-291">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-292">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-292">Office 2019 for Mac</span></span><br><span data-ttu-id="50980-293">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-293">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="50980-294">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-294">- TaskPane</span></span><br><span data-ttu-id="50980-295">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-295">
        - Content</span></span><br><span data-ttu-id="50980-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="50980-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="50980-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="50980-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="50980-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="50980-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="50980-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="50980-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="50980-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="50980-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-307">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-307">- BindingEvents</span></span><br><span data-ttu-id="50980-308">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-308">
        - CompressedFile</span></span><br><span data-ttu-id="50980-309">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-309">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-310">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-310">
        - File</span></span><br><span data-ttu-id="50980-311">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-311">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-312">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-312">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-313">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-313">
        - PdfFile</span></span><br><span data-ttu-id="50980-314">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-314">
        - Selection</span></span><br><span data-ttu-id="50980-315">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-315">
        - Settings</span></span><br><span data-ttu-id="50980-316">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-316">
        - TableBindings</span></span><br><span data-ttu-id="50980-317">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-317">
        - TableCoercion</span></span><br><span data-ttu-id="50980-318">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-318">
        - TextBindings</span></span><br><span data-ttu-id="50980-319">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-319">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-320">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-320">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="50980-321">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-321">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="50980-322">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-322">- TaskPane</span></span><br><span data-ttu-id="50980-323">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="50980-323">
        - Content</span></span></td>
    <td><span data-ttu-id="50980-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="50980-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="50980-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="50980-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-327">- BindingEvents</span></span><br><span data-ttu-id="50980-328">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-328">
        - CompressedFile</span></span><br><span data-ttu-id="50980-329">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-329">
        - DocumentEvents</span></span><br><span data-ttu-id="50980-330">
        - File</span><span class="sxs-lookup"><span data-stu-id="50980-330">
        - File</span></span><br><span data-ttu-id="50980-331">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-331">
        - MatrixBindings</span></span><br><span data-ttu-id="50980-332">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-332">
        - MatrixCoercion</span></span><br><span data-ttu-id="50980-333">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-333">
        - PdfFile</span></span><br><span data-ttu-id="50980-334">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-334">
        - Selection</span></span><br><span data-ttu-id="50980-335">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-335">
        - Settings</span></span><br><span data-ttu-id="50980-336">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-336">
        - TableBindings</span></span><br><span data-ttu-id="50980-337">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-337">
        - TableCoercion</span></span><br><span data-ttu-id="50980-338">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-338">
        - TextBindings</span></span><br><span data-ttu-id="50980-339">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-339">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="50980-340">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="50980-340">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="50980-341">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-341">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="50980-342">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-342">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="50980-343">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-343">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="50980-344">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-344">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="50980-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-346">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-346">Office on the web</span></span></td>
    <td><span data-ttu-id="50980-347">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-347">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="50980-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-349">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-349">Office on Windows</span></span><br><span data-ttu-id="50980-350">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-350">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="50980-351">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="50980-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-353">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-353">Office for Mac</span></span><br><span data-ttu-id="50980-354">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-354">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="50980-355">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="50980-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="50980-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="50980-357">Outlook</span><span class="sxs-lookup"><span data-stu-id="50980-357">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="50980-358">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-358">Platform</span></span></th>
    <th><span data-ttu-id="50980-359">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-359">Extension points</span></span></th>
    <th><span data-ttu-id="50980-360">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-360">API requirement sets</span></span></th>
    <th><span data-ttu-id="50980-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-362">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-362">Office on the web</span></span><br><span data-ttu-id="50980-363">(современная версия)</span><span class="sxs-lookup"><span data-stu-id="50980-363">Modern</span></span></td>
    <td> <span data-ttu-id="50980-364">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-364">- Mail Read</span></span><br><span data-ttu-id="50980-365">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-365">
      - Mail Compose</span></span><br><span data-ttu-id="50980-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="50980-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="50980-374">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-375">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-375">Office on the web</span></span><br><span data-ttu-id="50980-376">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="50980-376">(classic)</span></span></td>
    <td> <span data-ttu-id="50980-377">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-377">- Mail Read</span></span><br><span data-ttu-id="50980-378">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-378">
      - Mail Compose</span></span><br><span data-ttu-id="50980-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="50980-386">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-387">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-387">Office on Windows</span></span><br><span data-ttu-id="50980-388">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-389">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-389">- Mail Read</span></span><br><span data-ttu-id="50980-390">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-390">
      - Mail Compose</span></span><br><span data-ttu-id="50980-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="50980-392">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="50980-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="50980-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="50980-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="50980-400">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-400">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-401">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-401">Office 2019 on Windows</span></span><br><span data-ttu-id="50980-402">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-402">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-403">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-403">- Mail Read</span></span><br><span data-ttu-id="50980-404">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-404">
      - Mail Compose</span></span><br><span data-ttu-id="50980-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="50980-406">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="50980-406">
      - Modules</span></span></td>
    <td> <span data-ttu-id="50980-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="50980-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="50980-414">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-415">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-415">Office 2016 on Windows</span></span><br><span data-ttu-id="50980-416">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-416">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-417">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-417">- Mail Read</span></span><br><span data-ttu-id="50980-418">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-418">
      - Mail Compose</span></span><br><span data-ttu-id="50980-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="50980-420">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="50980-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="50980-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="50980-425">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-426">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-426">Office 2013 on Windows</span></span><br><span data-ttu-id="50980-427">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-427">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-428">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-428">- Mail Read</span></span><br><span data-ttu-id="50980-429">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-429">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="50980-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="50980-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="50980-434">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-434">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-435">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="50980-435">Office apps on iOS</span></span><br><span data-ttu-id="50980-436">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-436">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-437">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-437">- Mail Read</span></span><br><span data-ttu-id="50980-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="50980-444">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-445">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-445">Office apps on Mac</span></span><br><span data-ttu-id="50980-446">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-446">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-447">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-447">- Mail Read</span></span><br><span data-ttu-id="50980-448">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-448">
      - Mail Compose</span></span><br><span data-ttu-id="50980-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="50980-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="50980-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="50980-457">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-458">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-458">Office 2019 for Mac</span></span><br><span data-ttu-id="50980-459">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-460">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-460">- Mail Read</span></span><br><span data-ttu-id="50980-461">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-461">
      - Mail Compose</span></span><br><span data-ttu-id="50980-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="50980-469">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-470">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-470">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="50980-471">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-471">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-472">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-472">- Mail Read</span></span><br><span data-ttu-id="50980-473">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="50980-473">
      - Mail Compose</span></span><br><span data-ttu-id="50980-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="50980-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="50980-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="50980-481">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-482">Office для Android</span><span class="sxs-lookup"><span data-stu-id="50980-482">Office apps on Android</span></span><br><span data-ttu-id="50980-483">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-483">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-484">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="50980-484">- Mail Read</span></span><br><span data-ttu-id="50980-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="50980-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="50980-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="50980-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="50980-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="50980-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="50980-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="50980-491">Недоступно</span><span class="sxs-lookup"><span data-stu-id="50980-491">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="50980-492">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="50980-492">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="50980-493">Word</span><span class="sxs-lookup"><span data-stu-id="50980-493">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="50980-494">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-494">Platform</span></span></th>
    <th><span data-ttu-id="50980-495">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-495">Extension points</span></span></th>
    <th><span data-ttu-id="50980-496">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-496">API requirement sets</span></span></th>
    <th><span data-ttu-id="50980-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-498">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-498">Office on the web</span></span></td>
    <td> <span data-ttu-id="50980-499">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-499">- TaskPane</span></span><br><span data-ttu-id="50980-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="50980-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="50980-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="50980-507">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-507">- BindingEvents</span></span><br><span data-ttu-id="50980-508">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-508">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-509">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-509">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-510">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-510">
         - File</span></span><br><span data-ttu-id="50980-511">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-511">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-512">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-512">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-513">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-513">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-514">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-514">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-515">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-515">
         - PdfFile</span></span><br><span data-ttu-id="50980-516">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-516">
         - Selection</span></span><br><span data-ttu-id="50980-517">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-517">
         - Settings</span></span><br><span data-ttu-id="50980-518">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-518">
         - TableBindings</span></span><br><span data-ttu-id="50980-519">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-519">
         - TableCoercion</span></span><br><span data-ttu-id="50980-520">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-520">
         - TextBindings</span></span><br><span data-ttu-id="50980-521">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-521">
         - TextCoercion</span></span><br><span data-ttu-id="50980-522">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-522">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-523">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-523">Office on Windows</span></span><br><span data-ttu-id="50980-524">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-524">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-525">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-525">- TaskPane</span></span><br><span data-ttu-id="50980-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="50980-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="50980-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="50980-533">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-533">- BindingEvents</span></span><br><span data-ttu-id="50980-534">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-534">
         - CompressedFile</span></span><br><span data-ttu-id="50980-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-536">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-537">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-537">
         - File</span></span><br><span data-ttu-id="50980-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-539">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-542">
         - PdfFile</span></span><br><span data-ttu-id="50980-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-543">
         - Selection</span></span><br><span data-ttu-id="50980-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-544">
         - Settings</span></span><br><span data-ttu-id="50980-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-545">
         - TableBindings</span></span><br><span data-ttu-id="50980-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-546">
         - TableCoercion</span></span><br><span data-ttu-id="50980-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-547">
         - TextBindings</span></span><br><span data-ttu-id="50980-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-548">
         - TextCoercion</span></span><br><span data-ttu-id="50980-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-549">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-550">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-550">Office 2019 on Windows</span></span><br><span data-ttu-id="50980-551">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-551">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-552">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-552">- TaskPane</span></span><br><span data-ttu-id="50980-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="50980-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="50980-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-559">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-559">- BindingEvents</span></span><br><span data-ttu-id="50980-560">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-560">
         - CompressedFile</span></span><br><span data-ttu-id="50980-561">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-561">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-562">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-562">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-563">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-563">
         - File</span></span><br><span data-ttu-id="50980-564">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-564">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-565">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-565">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-566">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-566">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-567">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-567">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-568">
         - PdfFile</span></span><br><span data-ttu-id="50980-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-569">
         - Selection</span></span><br><span data-ttu-id="50980-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-570">
         - Settings</span></span><br><span data-ttu-id="50980-571">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-571">
         - TableBindings</span></span><br><span data-ttu-id="50980-572">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-572">
         - TableCoercion</span></span><br><span data-ttu-id="50980-573">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-573">
         - TextBindings</span></span><br><span data-ttu-id="50980-574">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-574">
         - TextCoercion</span></span><br><span data-ttu-id="50980-575">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-575">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-576">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-576">Office 2016 on Windows</span></span><br><span data-ttu-id="50980-577">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-577">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-578">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-578">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="50980-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-582">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-582">- BindingEvents</span></span><br><span data-ttu-id="50980-583">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-583">
         - CompressedFile</span></span><br><span data-ttu-id="50980-584">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-584">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-585">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-586">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-586">
         - File</span></span><br><span data-ttu-id="50980-587">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-587">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-588">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-588">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-589">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-589">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-590">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-590">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-591">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-591">
         - PdfFile</span></span><br><span data-ttu-id="50980-592">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-592">
         - Selection</span></span><br><span data-ttu-id="50980-593">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-593">
         - Settings</span></span><br><span data-ttu-id="50980-594">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-594">
         - TableBindings</span></span><br><span data-ttu-id="50980-595">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-595">
         - TableCoercion</span></span><br><span data-ttu-id="50980-596">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-596">
         - TextBindings</span></span><br><span data-ttu-id="50980-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-597">
         - TextCoercion</span></span><br><span data-ttu-id="50980-598">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-598">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-599">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-599">Office 2013 on Windows</span></span><br><span data-ttu-id="50980-600">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-600">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-601">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-601">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="50980-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="50980-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-604">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-604">- BindingEvents</span></span><br><span data-ttu-id="50980-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-605">
         - CompressedFile</span></span><br><span data-ttu-id="50980-606">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-606">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-607">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-608">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-608">
         - File</span></span><br><span data-ttu-id="50980-609">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-609">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-610">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-610">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-611">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-611">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-612">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-612">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-613">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-613">
         - PdfFile</span></span><br><span data-ttu-id="50980-614">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-614">
         - Selection</span></span><br><span data-ttu-id="50980-615">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-615">
         - Settings</span></span><br><span data-ttu-id="50980-616">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-616">
         - TableBindings</span></span><br><span data-ttu-id="50980-617">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-617">
         - TableCoercion</span></span><br><span data-ttu-id="50980-618">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-618">
         - TextBindings</span></span><br><span data-ttu-id="50980-619">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-619">
         - TextCoercion</span></span><br><span data-ttu-id="50980-620">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-620">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-621">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="50980-621">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="50980-622">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-622">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-623">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-623">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="50980-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="50980-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="50980-629">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-629">- BindingEvents</span></span><br><span data-ttu-id="50980-630">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-630">
         - CompressedFile</span></span><br><span data-ttu-id="50980-631">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-631">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-632">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-633">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-633">
         - File</span></span><br><span data-ttu-id="50980-634">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-634">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-635">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-635">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-636">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-636">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-637">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-637">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-638">
         - PdfFile</span></span><br><span data-ttu-id="50980-639">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-639">
         - Selection</span></span><br><span data-ttu-id="50980-640">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-640">
         - Settings</span></span><br><span data-ttu-id="50980-641">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-641">
         - TableBindings</span></span><br><span data-ttu-id="50980-642">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-642">
         - TableCoercion</span></span><br><span data-ttu-id="50980-643">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-643">
         - TextBindings</span></span><br><span data-ttu-id="50980-644">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-644">
         - TextCoercion</span></span><br><span data-ttu-id="50980-645">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-645">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-646">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-646">Office apps on Mac</span></span><br><span data-ttu-id="50980-647">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-647">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-648">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-648">- TaskPane</span></span><br><span data-ttu-id="50980-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="50980-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="50980-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="50980-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-656">- BindingEvents</span></span><br><span data-ttu-id="50980-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-657">
         - CompressedFile</span></span><br><span data-ttu-id="50980-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-659">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-660">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-660">
         - File</span></span><br><span data-ttu-id="50980-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-662">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-665">
         - PdfFile</span></span><br><span data-ttu-id="50980-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-666">
         - Selection</span></span><br><span data-ttu-id="50980-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-667">
         - Settings</span></span><br><span data-ttu-id="50980-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-668">
         - TableBindings</span></span><br><span data-ttu-id="50980-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-669">
         - TableCoercion</span></span><br><span data-ttu-id="50980-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-670">
         - TextBindings</span></span><br><span data-ttu-id="50980-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-671">
         - TextCoercion</span></span><br><span data-ttu-id="50980-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-673">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-673">Office 2019 for Mac</span></span><br><span data-ttu-id="50980-674">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-674">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-675">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-675">- TaskPane</span></span><br><span data-ttu-id="50980-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="50980-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="50980-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="50980-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="50980-682">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-682">- BindingEvents</span></span><br><span data-ttu-id="50980-683">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-683">
         - CompressedFile</span></span><br><span data-ttu-id="50980-684">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-684">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-685">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-685">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-686">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-686">
         - File</span></span><br><span data-ttu-id="50980-687">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-687">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-688">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-688">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-689">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-689">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-690">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-690">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-691">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-691">
         - PdfFile</span></span><br><span data-ttu-id="50980-692">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-692">
         - Selection</span></span><br><span data-ttu-id="50980-693">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-693">
         - Settings</span></span><br><span data-ttu-id="50980-694">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-694">
         - TableBindings</span></span><br><span data-ttu-id="50980-695">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-695">
         - TableCoercion</span></span><br><span data-ttu-id="50980-696">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-696">
         - TextBindings</span></span><br><span data-ttu-id="50980-697">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-697">
         - TextCoercion</span></span><br><span data-ttu-id="50980-698">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-698">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-699">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-699">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="50980-700">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-700">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-701">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-701">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="50980-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="50980-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="50980-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-705">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="50980-705">- BindingEvents</span></span><br><span data-ttu-id="50980-706">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-706">
         - CompressedFile</span></span><br><span data-ttu-id="50980-707">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="50980-707">
         - CustomXmlParts</span></span><br><span data-ttu-id="50980-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-708">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-709">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="50980-709">
         - File</span></span><br><span data-ttu-id="50980-710">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-710">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-711">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="50980-711">
         - MatrixBindings</span></span><br><span data-ttu-id="50980-712">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-712">
         - MatrixCoercion</span></span><br><span data-ttu-id="50980-713">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-713">
         - OoxmlCoercion</span></span><br><span data-ttu-id="50980-714">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-714">
         - PdfFile</span></span><br><span data-ttu-id="50980-715">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-715">
         - Selection</span></span><br><span data-ttu-id="50980-716">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="50980-716">
         - Settings</span></span><br><span data-ttu-id="50980-717">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="50980-717">
         - TableBindings</span></span><br><span data-ttu-id="50980-718">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-718">
         - TableCoercion</span></span><br><span data-ttu-id="50980-719">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="50980-719">
         - TextBindings</span></span><br><span data-ttu-id="50980-720">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-720">
         - TextCoercion</span></span><br><span data-ttu-id="50980-721">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="50980-721">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="50980-722">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="50980-722">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="50980-723">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="50980-723">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="50980-724">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-724">Platform</span></span></th>
    <th><span data-ttu-id="50980-725">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-725">Extension points</span></span></th>
    <th><span data-ttu-id="50980-726">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-726">API requirement sets</span></span></th>
    <th><span data-ttu-id="50980-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-728">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-728">Office on the web</span></span></td>
    <td> <span data-ttu-id="50980-729">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-729">- Content</span></span><br><span data-ttu-id="50980-730">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-730">
         - TaskPane</span></span><br><span data-ttu-id="50980-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="50980-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="50980-736">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-736">- ActiveView</span></span><br><span data-ttu-id="50980-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-737">
         - CompressedFile</span></span><br><span data-ttu-id="50980-738">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-738">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-739">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-739">
         - File</span></span><br><span data-ttu-id="50980-740">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-740">
         - PdfFile</span></span><br><span data-ttu-id="50980-741">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-741">
         - Selection</span></span><br><span data-ttu-id="50980-742">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-742">
         - Settings</span></span><br><span data-ttu-id="50980-743">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-743">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-744">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-744">Office on Windows</span></span><br><span data-ttu-id="50980-745">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-745">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-746">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-746">- Content</span></span><br><span data-ttu-id="50980-747">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-747">
         - TaskPane</span></span><br><span data-ttu-id="50980-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="50980-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="50980-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-753">- ActiveView</span></span><br><span data-ttu-id="50980-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-754">
         - CompressedFile</span></span><br><span data-ttu-id="50980-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-755">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-756">
         - File</span></span><br><span data-ttu-id="50980-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-757">
         - PdfFile</span></span><br><span data-ttu-id="50980-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-758">
         - Selection</span></span><br><span data-ttu-id="50980-759">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-759">
         - Settings</span></span><br><span data-ttu-id="50980-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-761">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-761">Office 2019 on Windows</span></span><br><span data-ttu-id="50980-762">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-763">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-763">- Content</span></span><br><span data-ttu-id="50980-764">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-764">
         - TaskPane</span></span><br><span data-ttu-id="50980-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-768">- ActiveView</span></span><br><span data-ttu-id="50980-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-769">
         - CompressedFile</span></span><br><span data-ttu-id="50980-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-770">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-771">
         - File</span></span><br><span data-ttu-id="50980-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-772">
         - PdfFile</span></span><br><span data-ttu-id="50980-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-773">
         - Selection</span></span><br><span data-ttu-id="50980-774">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-774">
         - Settings</span></span><br><span data-ttu-id="50980-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-776">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-776">Office 2016 on Windows</span></span><br><span data-ttu-id="50980-777">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-778">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-778">- Content</span></span><br><span data-ttu-id="50980-779">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="50980-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="50980-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-782">- ActiveView</span></span><br><span data-ttu-id="50980-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-783">
         - CompressedFile</span></span><br><span data-ttu-id="50980-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-784">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-785">
         - File</span></span><br><span data-ttu-id="50980-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-786">
         - PdfFile</span></span><br><span data-ttu-id="50980-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-787">
         - Selection</span></span><br><span data-ttu-id="50980-788">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-788">
         - Settings</span></span><br><span data-ttu-id="50980-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-790">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-790">Office 2013 on Windows</span></span><br><span data-ttu-id="50980-791">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-792">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-792">- Content</span></span><br><span data-ttu-id="50980-793">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="50980-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="50980-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="50980-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-796">- ActiveView</span></span><br><span data-ttu-id="50980-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-797">
         - CompressedFile</span></span><br><span data-ttu-id="50980-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-798">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-799">
         - File</span></span><br><span data-ttu-id="50980-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-800">
         - PdfFile</span></span><br><span data-ttu-id="50980-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-801">
         - Selection</span></span><br><span data-ttu-id="50980-802">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-802">
         - Settings</span></span><br><span data-ttu-id="50980-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-804">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="50980-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="50980-805">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-806">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-806">- Content</span></span><br><span data-ttu-id="50980-807">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="50980-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-811">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-811">- ActiveView</span></span><br><span data-ttu-id="50980-812">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-812">
         - CompressedFile</span></span><br><span data-ttu-id="50980-813">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-813">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-814">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-814">
         - File</span></span><br><span data-ttu-id="50980-815">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-815">
         - PdfFile</span></span><br><span data-ttu-id="50980-816">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-816">
         - Selection</span></span><br><span data-ttu-id="50980-817">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-817">
         - Settings</span></span><br><span data-ttu-id="50980-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-818">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-819">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-819">Office apps on Mac</span></span><br><span data-ttu-id="50980-820">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="50980-820">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="50980-821">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-821">- Content</span></span><br><span data-ttu-id="50980-822">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-822">
         - TaskPane</span></span><br><span data-ttu-id="50980-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="50980-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="50980-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="50980-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="50980-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-828">- ActiveView</span></span><br><span data-ttu-id="50980-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-829">
         - CompressedFile</span></span><br><span data-ttu-id="50980-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-830">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-831">
         - File</span></span><br><span data-ttu-id="50980-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-832">
         - PdfFile</span></span><br><span data-ttu-id="50980-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-833">
         - Selection</span></span><br><span data-ttu-id="50980-834">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-834">
         - Settings</span></span><br><span data-ttu-id="50980-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-836">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-836">Office 2019 for Mac</span></span><br><span data-ttu-id="50980-837">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-837">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-838">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-838">- Content</span></span><br><span data-ttu-id="50980-839">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-839">
         - TaskPane</span></span><br><span data-ttu-id="50980-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-843">- ActiveView</span></span><br><span data-ttu-id="50980-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-844">
         - CompressedFile</span></span><br><span data-ttu-id="50980-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-845">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-846">
         - File</span></span><br><span data-ttu-id="50980-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-847">
         - PdfFile</span></span><br><span data-ttu-id="50980-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-848">
         - Selection</span></span><br><span data-ttu-id="50980-849">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-849">
         - Settings</span></span><br><span data-ttu-id="50980-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-851">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-851">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="50980-852">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-852">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-853">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-853">- Content</span></span><br><span data-ttu-id="50980-854">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-854">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="50980-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="50980-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-857">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="50980-857">- ActiveView</span></span><br><span data-ttu-id="50980-858">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="50980-858">
         - CompressedFile</span></span><br><span data-ttu-id="50980-859">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-859">
         - DocumentEvents</span></span><br><span data-ttu-id="50980-860">
         - File</span><span class="sxs-lookup"><span data-stu-id="50980-860">
         - File</span></span><br><span data-ttu-id="50980-861">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="50980-861">
         - PdfFile</span></span><br><span data-ttu-id="50980-862">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="50980-862">
         - Selection</span></span><br><span data-ttu-id="50980-863">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-863">
         - Settings</span></span><br><span data-ttu-id="50980-864">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-864">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="50980-865">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="50980-865">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="50980-866">OneNote</span><span class="sxs-lookup"><span data-stu-id="50980-866">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="50980-867">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-867">Platform</span></span></th>
    <th><span data-ttu-id="50980-868">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-868">Extension points</span></span></th>
    <th><span data-ttu-id="50980-869">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-869">API requirement sets</span></span></th>
    <th><span data-ttu-id="50980-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-871">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="50980-871">Office on the web</span></span></td>
    <td> <span data-ttu-id="50980-872">- Контент</span><span class="sxs-lookup"><span data-stu-id="50980-872">- Content</span></span><br><span data-ttu-id="50980-873">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-873">
         - TaskPane</span></span><br><span data-ttu-id="50980-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="50980-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="50980-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="50980-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="50980-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-878">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="50980-878">- DocumentEvents</span></span><br><span data-ttu-id="50980-879">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-879">
         - HtmlCoercion</span></span><br><span data-ttu-id="50980-880">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="50980-880">
         - Settings</span></span><br><span data-ttu-id="50980-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="50980-882">Project</span><span class="sxs-lookup"><span data-stu-id="50980-882">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="50980-883">Платформа</span><span class="sxs-lookup"><span data-stu-id="50980-883">Platform</span></span></th>
    <th><span data-ttu-id="50980-884">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="50980-884">Extension points</span></span></th>
    <th><span data-ttu-id="50980-885">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="50980-885">API requirement sets</span></span></th>
    <th><span data-ttu-id="50980-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="50980-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-887">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-887">Office 2019 on Windows</span></span><br><span data-ttu-id="50980-888">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-888">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-889">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-889">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-891">- Selection</span><span class="sxs-lookup"><span data-stu-id="50980-891">- Selection</span></span><br><span data-ttu-id="50980-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-892">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-893">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-893">Office 2016 on Windows</span></span><br><span data-ttu-id="50980-894">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-894">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-895">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-895">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-897">- Selection</span><span class="sxs-lookup"><span data-stu-id="50980-897">- Selection</span></span><br><span data-ttu-id="50980-898">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-898">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="50980-899">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="50980-899">Office 2013 on Windows</span></span><br><span data-ttu-id="50980-900">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="50980-900">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="50980-901">- Область задач</span><span class="sxs-lookup"><span data-stu-id="50980-901">- TaskPane</span></span></td>
    <td> <span data-ttu-id="50980-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="50980-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="50980-903">- Selection</span><span class="sxs-lookup"><span data-stu-id="50980-903">- Selection</span></span><br><span data-ttu-id="50980-904">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="50980-904">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="50980-905">См. также</span><span class="sxs-lookup"><span data-stu-id="50980-905">See also</span></span>

- [<span data-ttu-id="50980-906">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="50980-906">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="50980-907">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="50980-907">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="50980-908">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="50980-908">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="50980-909">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="50980-909">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="50980-910">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="50980-910">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="50980-911">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="50980-911">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="50980-912">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="50980-912">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="50980-913">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="50980-913">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="50980-914">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="50980-914">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="50980-915">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="50980-915">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="50980-916">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="50980-916">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
