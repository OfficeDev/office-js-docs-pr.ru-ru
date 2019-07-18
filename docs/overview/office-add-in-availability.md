---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: 2bfeb7cc5c6e8846f1d882abf3a0149302e53914
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771837"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="3e459-103">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="3e459-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="3e459-104">Работа надстройки Office может зависеть от ведущего приложения Office, набора требований, элемента или версии API.</span><span class="sxs-lookup"><span data-stu-id="3e459-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="3e459-105">В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, которые в настоящее время поддерживаются для всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="3e459-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="3e459-106">Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API.</span><span class="sxs-lookup"><span data-stu-id="3e459-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="3e459-107">Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also).</span><span class="sxs-lookup"><span data-stu-id="3e459-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="3e459-108">Excel</span><span class="sxs-lookup"><span data-stu-id="3e459-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3e459-109">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3e459-110">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3e459-111">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3e459-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-113">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="3e459-114">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-114">- TaskPane</span></span><br><span data-ttu-id="3e459-115">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-115">
        - Content</span></span><br><span data-ttu-id="3e459-116">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-116">
        - Custom Functions</span></span><br><span data-ttu-id="3e459-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="3e459-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3e459-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3e459-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3e459-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3e459-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3e459-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3e459-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3e459-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3e459-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3e459-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3e459-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3e459-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3e459-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-130">
        - BindingEvents</span></span><br><span data-ttu-id="3e459-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-131">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-132">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-133">
        - File</span></span><br><span data-ttu-id="3e459-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-134">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-136">
        - Selection</span></span><br><span data-ttu-id="3e459-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-137">
        - Settings</span></span><br><span data-ttu-id="3e459-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-138">
        - TableBindings</span></span><br><span data-ttu-id="3e459-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-139">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-140">
        - TextBindings</span></span><br><span data-ttu-id="3e459-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-142">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-142">Office on Windows</span></span><br><span data-ttu-id="3e459-143">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-144">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-144">- TaskPane</span></span><br><span data-ttu-id="3e459-145">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-145">
        - Content</span></span><br><span data-ttu-id="3e459-146">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-146">
        - Custom Functions</span></span><br><span data-ttu-id="3e459-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстройки</a>
    </span><span class="sxs-lookup"><span data-stu-id="3e459-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3e459-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3e459-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3e459-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3e459-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3e459-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3e459-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3e459-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3e459-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3e459-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3e459-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3e459-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3e459-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-160">
        - BindingEvents</span></span><br><span data-ttu-id="3e459-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-161">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-162">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-163">
        - File</span></span><br><span data-ttu-id="3e459-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-164">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-166">
        - Selection</span></span><br><span data-ttu-id="3e459-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-167">
        - Settings</span></span><br><span data-ttu-id="3e459-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-168">
        - TableBindings</span></span><br><span data-ttu-id="3e459-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-169">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-170">
        - TextBindings</span></span><br><span data-ttu-id="3e459-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-172">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-172">Office 2019 on Windows</span></span><br><span data-ttu-id="3e459-173">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3e459-174">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-174">- TaskPane</span></span><br><span data-ttu-id="3e459-175">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-175">
        - Content</span></span><br><span data-ttu-id="3e459-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3e459-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3e459-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3e459-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3e459-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3e459-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3e459-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3e459-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3e459-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3e459-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3e459-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-187">- BindingEvents</span></span><br><span data-ttu-id="3e459-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-188">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-189">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-190">
        - File</span></span><br><span data-ttu-id="3e459-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-191">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-193">
        - Selection</span></span><br><span data-ttu-id="3e459-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-194">
        - Settings</span></span><br><span data-ttu-id="3e459-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-195">
        - TableBindings</span></span><br><span data-ttu-id="3e459-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-196">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-197">
        - TextBindings</span></span><br><span data-ttu-id="3e459-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-199">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-199">Office 2016 on Windows</span></span><br><span data-ttu-id="3e459-200">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3e459-201">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-201">- TaskPane</span></span><br><span data-ttu-id="3e459-202">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-202">
        - Content</span></span></td>
    <td><span data-ttu-id="3e459-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3e459-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3e459-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-206">- BindingEvents</span></span><br><span data-ttu-id="3e459-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-207">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-208">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-209">
        - File</span></span><br><span data-ttu-id="3e459-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-210">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-212">
        - Selection</span></span><br><span data-ttu-id="3e459-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-213">
        - Settings</span></span><br><span data-ttu-id="3e459-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-214">
        - TableBindings</span></span><br><span data-ttu-id="3e459-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-215">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-216">
        - TextBindings</span></span><br><span data-ttu-id="3e459-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-218">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-218">Office 2013 on Windows</span></span><br><span data-ttu-id="3e459-219">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3e459-220">
        - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-220">
        - TaskPane</span></span><br><span data-ttu-id="3e459-221">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="3e459-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3e459-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3e459-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3e459-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-224">
        - BindingEvents</span></span><br><span data-ttu-id="3e459-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-225">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-226">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-227">
        - File</span></span><br><span data-ttu-id="3e459-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-228">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-230">
        - Selection</span></span><br><span data-ttu-id="3e459-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-231">
        - Settings</span></span><br><span data-ttu-id="3e459-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-232">
        - TableBindings</span></span><br><span data-ttu-id="3e459-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-233">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-234">
        - TextBindings</span></span><br><span data-ttu-id="3e459-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-236">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="3e459-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3e459-237">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3e459-238">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-238">- TaskPane</span></span><br><span data-ttu-id="3e459-239">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-239">
        - Content</span></span><br><span data-ttu-id="3e459-240">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3e459-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3e459-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3e459-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3e459-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3e459-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3e459-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3e459-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3e459-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3e459-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3e459-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3e459-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3e459-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-252">- BindingEvents</span></span><br><span data-ttu-id="3e459-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-253">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-254">
        - File</span></span><br><span data-ttu-id="3e459-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-255">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-257">
        - Selection</span></span><br><span data-ttu-id="3e459-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-258">
        - Settings</span></span><br><span data-ttu-id="3e459-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-259">
        - TableBindings</span></span><br><span data-ttu-id="3e459-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-260">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-261">
        - TextBindings</span></span><br><span data-ttu-id="3e459-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-263">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-263">Office apps on Mac</span></span><br><span data-ttu-id="3e459-264">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3e459-265">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-265">- TaskPane</span></span><br><span data-ttu-id="3e459-266">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-266">
        - Content</span></span><br><span data-ttu-id="3e459-267">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-267">
        - Custom Functions</span></span><br><span data-ttu-id="3e459-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3e459-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3e459-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3e459-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3e459-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3e459-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3e459-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3e459-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3e459-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3e459-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3e459-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3e459-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3e459-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-281">- BindingEvents</span></span><br><span data-ttu-id="3e459-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-282">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-283">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-284">
        - File</span></span><br><span data-ttu-id="3e459-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-285">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-287">
        - PdfFile</span></span><br><span data-ttu-id="3e459-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-288">
        - Selection</span></span><br><span data-ttu-id="3e459-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-289">
        - Settings</span></span><br><span data-ttu-id="3e459-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-290">
        - TableBindings</span></span><br><span data-ttu-id="3e459-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-291">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-292">
        - TextBindings</span></span><br><span data-ttu-id="3e459-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-294">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-294">Office 2019 for Mac</span></span><br><span data-ttu-id="3e459-295">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3e459-296">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-296">- TaskPane</span></span><br><span data-ttu-id="3e459-297">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-297">
        - Content</span></span><br><span data-ttu-id="3e459-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3e459-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3e459-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3e459-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3e459-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3e459-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3e459-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3e459-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3e459-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3e459-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3e459-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-309">- BindingEvents</span></span><br><span data-ttu-id="3e459-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-310">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-311">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-312">
        - File</span></span><br><span data-ttu-id="3e459-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-313">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-315">
        - PdfFile</span></span><br><span data-ttu-id="3e459-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-316">
        - Selection</span></span><br><span data-ttu-id="3e459-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-317">
        - Settings</span></span><br><span data-ttu-id="3e459-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-318">
        - TableBindings</span></span><br><span data-ttu-id="3e459-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-319">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-320">
        - TextBindings</span></span><br><span data-ttu-id="3e459-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-322">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3e459-323">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3e459-324">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-324">- TaskPane</span></span><br><span data-ttu-id="3e459-325">
        - Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-325">
        - Content</span></span></td>
    <td><span data-ttu-id="3e459-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3e459-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3e459-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3e459-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-329">- BindingEvents</span></span><br><span data-ttu-id="3e459-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-330">
        - CompressedFile</span></span><br><span data-ttu-id="3e459-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-331">
        - DocumentEvents</span></span><br><span data-ttu-id="3e459-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="3e459-332">
        - File</span></span><br><span data-ttu-id="3e459-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-333">
        - MatrixBindings</span></span><br><span data-ttu-id="3e459-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="3e459-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-335">
        - PdfFile</span></span><br><span data-ttu-id="3e459-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-336">
        - Selection</span></span><br><span data-ttu-id="3e459-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-337">
        - Settings</span></span><br><span data-ttu-id="3e459-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-338">
        - TableBindings</span></span><br><span data-ttu-id="3e459-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-339">
        - TableCoercion</span></span><br><span data-ttu-id="3e459-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-340">
        - TextBindings</span></span><br><span data-ttu-id="3e459-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="3e459-342">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="3e459-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="3e459-343">Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3e459-344">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3e459-345">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3e459-346">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3e459-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-348">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-348">Office on the web</span></span></td>
    <td><span data-ttu-id="3e459-349">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3e459-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-351">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-351">Office on Windows</span></span><br><span data-ttu-id="3e459-352">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3e459-353">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3e459-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-355">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-355">Office for Mac</span></span><br><span data-ttu-id="3e459-356">(подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="3e459-357">
        - Пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="3e459-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3e459-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="3e459-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="3e459-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3e459-360">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-360">Platform</span></span></th>
    <th><span data-ttu-id="3e459-361">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-361">Extension points</span></span></th>
    <th><span data-ttu-id="3e459-362">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="3e459-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-364">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-364">Office on the web</span></span><br><span data-ttu-id="3e459-365">(новый)</span><span class="sxs-lookup"><span data-stu-id="3e459-365">New</span></span></td>
    <td> <span data-ttu-id="3e459-366">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-366">- Mail Read</span></span><br><span data-ttu-id="3e459-367">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-367">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3e459-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3e459-376">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-377">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-377">Office on the web</span></span><br><span data-ttu-id="3e459-378">(классическая версия)</span><span class="sxs-lookup"><span data-stu-id="3e459-378">(classic)</span></span></td>
    <td> <span data-ttu-id="3e459-379">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-379">- Mail Read</span></span><br><span data-ttu-id="3e459-380">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-380">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3e459-388">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-389">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-389">Office on Windows</span></span><br><span data-ttu-id="3e459-390">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-391">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-391">- Mail Read</span></span><br><span data-ttu-id="3e459-392">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-392">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3e459-394">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="3e459-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3e459-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3e459-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3e459-402">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-403">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-403">Office 2019 on Windows</span></span><br><span data-ttu-id="3e459-404">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-405">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-405">- Mail Read</span></span><br><span data-ttu-id="3e459-406">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-406">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3e459-408">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="3e459-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3e459-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3e459-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3e459-416">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-417">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-417">Office 2016 on Windows</span></span><br><span data-ttu-id="3e459-418">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-419">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-419">- Mail Read</span></span><br><span data-ttu-id="3e459-420">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-420">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3e459-422">
      - Модули</span><span class="sxs-lookup"><span data-stu-id="3e459-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3e459-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="3e459-427">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-428">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-428">Office 2013 on Windows</span></span><br><span data-ttu-id="3e459-429">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-430">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-430">- Mail Read</span></span><br><span data-ttu-id="3e459-431">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="3e459-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="3e459-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="3e459-436">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-437">Office для iOS</span><span class="sxs-lookup"><span data-stu-id="3e459-437">Office apps on iOS</span></span><br><span data-ttu-id="3e459-438">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-439">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-439">- Mail Read</span></span><br><span data-ttu-id="3e459-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="3e459-446">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-447">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-447">Office apps on Mac</span></span><br><span data-ttu-id="3e459-448">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-449">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-449">- Mail Read</span></span><br><span data-ttu-id="3e459-450">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-450">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3e459-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3e459-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3e459-459">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-460">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-460">Office 2019 for Mac</span></span><br><span data-ttu-id="3e459-461">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-462">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-462">- Mail Read</span></span><br><span data-ttu-id="3e459-463">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-463">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3e459-471">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-472">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3e459-473">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-474">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-474">- Mail Read</span></span><br><span data-ttu-id="3e459-475">
      - Создание сообщения почты</span><span class="sxs-lookup"><span data-stu-id="3e459-475">
      - Mail Compose</span></span><br><span data-ttu-id="3e459-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3e459-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3e459-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3e459-483">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-484">Office для Android</span><span class="sxs-lookup"><span data-stu-id="3e459-484">Office apps on Android</span></span><br><span data-ttu-id="3e459-485">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-486">- Чтение почты</span><span class="sxs-lookup"><span data-stu-id="3e459-486">- Mail Read</span></span><br><span data-ttu-id="3e459-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3e459-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3e459-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3e459-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3e459-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3e459-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3e459-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="3e459-493">Недоступно</span><span class="sxs-lookup"><span data-stu-id="3e459-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="3e459-494">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="3e459-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="3e459-495">Word</span><span class="sxs-lookup"><span data-stu-id="3e459-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3e459-496">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-496">Platform</span></span></th>
    <th><span data-ttu-id="3e459-497">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-497">Extension points</span></span></th>
    <th><span data-ttu-id="3e459-498">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="3e459-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-500">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="3e459-501">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-501">- TaskPane</span></span><br><span data-ttu-id="3e459-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="3e459-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="3e459-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3e459-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-509">- BindingEvents</span></span><br><span data-ttu-id="3e459-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-511">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-512">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-512">
         - File</span></span><br><span data-ttu-id="3e459-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-514">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-517">
         - PdfFile</span></span><br><span data-ttu-id="3e459-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-518">
         - Selection</span></span><br><span data-ttu-id="3e459-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-519">
         - Settings</span></span><br><span data-ttu-id="3e459-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-520">
         - TableBindings</span></span><br><span data-ttu-id="3e459-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-521">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-522">
         - TextBindings</span></span><br><span data-ttu-id="3e459-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-523">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-525">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-525">Office on Windows</span></span><br><span data-ttu-id="3e459-526">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-527">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-527">- TaskPane</span></span><br><span data-ttu-id="3e459-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="3e459-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="3e459-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3e459-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-535">- BindingEvents</span></span><br><span data-ttu-id="3e459-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-536">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-538">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-539">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-539">
         - File</span></span><br><span data-ttu-id="3e459-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-541">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-544">
         - PdfFile</span></span><br><span data-ttu-id="3e459-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-545">
         - Selection</span></span><br><span data-ttu-id="3e459-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-546">
         - Settings</span></span><br><span data-ttu-id="3e459-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-547">
         - TableBindings</span></span><br><span data-ttu-id="3e459-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-548">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-549">
         - TextBindings</span></span><br><span data-ttu-id="3e459-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-550">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-552">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-552">Office 2019 on Windows</span></span><br><span data-ttu-id="3e459-553">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-554">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-554">- TaskPane</span></span><br><span data-ttu-id="3e459-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="3e459-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="3e459-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-561">- BindingEvents</span></span><br><span data-ttu-id="3e459-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-562">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-564">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-565">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-565">
         - File</span></span><br><span data-ttu-id="3e459-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-567">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-570">
         - PdfFile</span></span><br><span data-ttu-id="3e459-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-571">
         - Selection</span></span><br><span data-ttu-id="3e459-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-572">
         - Settings</span></span><br><span data-ttu-id="3e459-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-573">
         - TableBindings</span></span><br><span data-ttu-id="3e459-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-574">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-575">
         - TextBindings</span></span><br><span data-ttu-id="3e459-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-576">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-578">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-578">Office 2016 on Windows</span></span><br><span data-ttu-id="3e459-579">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-580">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3e459-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-584">- BindingEvents</span></span><br><span data-ttu-id="3e459-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-585">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-587">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-588">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-588">
         - File</span></span><br><span data-ttu-id="3e459-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-590">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-593">
         - PdfFile</span></span><br><span data-ttu-id="3e459-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-594">
         - Selection</span></span><br><span data-ttu-id="3e459-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-595">
         - Settings</span></span><br><span data-ttu-id="3e459-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-596">
         - TableBindings</span></span><br><span data-ttu-id="3e459-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-597">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-598">
         - TextBindings</span></span><br><span data-ttu-id="3e459-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-599">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-601">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-601">Office 2013 on Windows</span></span><br><span data-ttu-id="3e459-602">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-603">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3e459-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3e459-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-606">- BindingEvents</span></span><br><span data-ttu-id="3e459-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-607">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-609">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-610">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-610">
         - File</span></span><br><span data-ttu-id="3e459-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-612">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-615">
         - PdfFile</span></span><br><span data-ttu-id="3e459-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-616">
         - Selection</span></span><br><span data-ttu-id="3e459-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-617">
         - Settings</span></span><br><span data-ttu-id="3e459-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-618">
         - TableBindings</span></span><br><span data-ttu-id="3e459-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-619">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-620">
         - TextBindings</span></span><br><span data-ttu-id="3e459-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-621">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-623">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="3e459-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3e459-624">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-625">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="3e459-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="3e459-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="3e459-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-631">- BindingEvents</span></span><br><span data-ttu-id="3e459-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-632">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-634">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-635">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-635">
         - File</span></span><br><span data-ttu-id="3e459-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-637">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-640">
         - PdfFile</span></span><br><span data-ttu-id="3e459-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-641">
         - Selection</span></span><br><span data-ttu-id="3e459-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-642">
         - Settings</span></span><br><span data-ttu-id="3e459-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-643">
         - TableBindings</span></span><br><span data-ttu-id="3e459-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-644">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-645">
         - TextBindings</span></span><br><span data-ttu-id="3e459-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-646">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-648">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-648">Office apps on Mac</span></span><br><span data-ttu-id="3e459-649">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-650">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-650">- TaskPane</span></span><br><span data-ttu-id="3e459-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="3e459-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="3e459-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="3e459-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-658">- BindingEvents</span></span><br><span data-ttu-id="3e459-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-659">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-661">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-662">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-662">
         - File</span></span><br><span data-ttu-id="3e459-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-664">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-667">
         - PdfFile</span></span><br><span data-ttu-id="3e459-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-668">
         - Selection</span></span><br><span data-ttu-id="3e459-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-669">
         - Settings</span></span><br><span data-ttu-id="3e459-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-670">
         - TableBindings</span></span><br><span data-ttu-id="3e459-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-671">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-672">
         - TextBindings</span></span><br><span data-ttu-id="3e459-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-673">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-675">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-675">Office 2019 for Mac</span></span><br><span data-ttu-id="3e459-676">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-677">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-677">- TaskPane</span></span><br><span data-ttu-id="3e459-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="3e459-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3e459-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="3e459-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="3e459-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-684">- BindingEvents</span></span><br><span data-ttu-id="3e459-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-685">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-687">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-688">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-688">
         - File</span></span><br><span data-ttu-id="3e459-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-690">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-693">
         - PdfFile</span></span><br><span data-ttu-id="3e459-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-694">
         - Selection</span></span><br><span data-ttu-id="3e459-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-695">
         - Settings</span></span><br><span data-ttu-id="3e459-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-696">
         - TableBindings</span></span><br><span data-ttu-id="3e459-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-697">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-698">
         - TextBindings</span></span><br><span data-ttu-id="3e459-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-699">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-701">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3e459-702">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-703">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="3e459-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3e459-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3e459-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-707">- BindingEvents</span></span><br><span data-ttu-id="3e459-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-708">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3e459-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="3e459-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-710">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-711">
         - Файл</span><span class="sxs-lookup"><span data-stu-id="3e459-711">
         - File</span></span><br><span data-ttu-id="3e459-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-713">
         - MatrixBindings</span></span><br><span data-ttu-id="3e459-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="3e459-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3e459-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-716">
         - PdfFile</span></span><br><span data-ttu-id="3e459-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-717">
         - Selection</span></span><br><span data-ttu-id="3e459-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="3e459-718">
         - Settings</span></span><br><span data-ttu-id="3e459-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-719">
         - TableBindings</span></span><br><span data-ttu-id="3e459-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-720">
         - TableCoercion</span></span><br><span data-ttu-id="3e459-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3e459-721">
         - TextBindings</span></span><br><span data-ttu-id="3e459-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-722">
         - TextCoercion</span></span><br><span data-ttu-id="3e459-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3e459-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="3e459-724">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="3e459-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="3e459-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="3e459-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3e459-726">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-726">Platform</span></span></th>
    <th><span data-ttu-id="3e459-727">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-727">Extension points</span></span></th>
    <th><span data-ttu-id="3e459-728">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="3e459-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-730">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="3e459-731">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-731">- Content</span></span><br><span data-ttu-id="3e459-732">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-732">
         - TaskPane</span></span><br><span data-ttu-id="3e459-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3e459-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-737">- ActiveView</span></span><br><span data-ttu-id="3e459-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-738">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-739">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-740">
         - File</span></span><br><span data-ttu-id="3e459-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-741">
         - PdfFile</span></span><br><span data-ttu-id="3e459-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-742">
         - Selection</span></span><br><span data-ttu-id="3e459-743">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-743">
         - Settings</span></span><br><span data-ttu-id="3e459-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-745">Office для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-745">Office on Windows</span></span><br><span data-ttu-id="3e459-746">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-747">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-747">- Content</span></span><br><span data-ttu-id="3e459-748">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-748">
         - TaskPane</span></span><br><span data-ttu-id="3e459-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3e459-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-753">- ActiveView</span></span><br><span data-ttu-id="3e459-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-754">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-755">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-756">
         - File</span></span><br><span data-ttu-id="3e459-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-757">
         - PdfFile</span></span><br><span data-ttu-id="3e459-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-758">
         - Selection</span></span><br><span data-ttu-id="3e459-759">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-759">
         - Settings</span></span><br><span data-ttu-id="3e459-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-761">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-761">Office 2019 on Windows</span></span><br><span data-ttu-id="3e459-762">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-763">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-763">- Content</span></span><br><span data-ttu-id="3e459-764">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-764">
         - TaskPane</span></span><br><span data-ttu-id="3e459-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-768">- ActiveView</span></span><br><span data-ttu-id="3e459-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-769">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-770">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-771">
         - File</span></span><br><span data-ttu-id="3e459-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-772">
         - PdfFile</span></span><br><span data-ttu-id="3e459-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-773">
         - Selection</span></span><br><span data-ttu-id="3e459-774">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-774">
         - Settings</span></span><br><span data-ttu-id="3e459-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-776">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-776">Office 2016 on Windows</span></span><br><span data-ttu-id="3e459-777">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-778">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-778">- Content</span></span><br><span data-ttu-id="3e459-779">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3e459-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3e459-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-782">- ActiveView</span></span><br><span data-ttu-id="3e459-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-783">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-784">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-785">
         - File</span></span><br><span data-ttu-id="3e459-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-786">
         - PdfFile</span></span><br><span data-ttu-id="3e459-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-787">
         - Selection</span></span><br><span data-ttu-id="3e459-788">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-788">
         - Settings</span></span><br><span data-ttu-id="3e459-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-790">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-790">Office 2013 on Windows</span></span><br><span data-ttu-id="3e459-791">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-792">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-792">- Content</span></span><br><span data-ttu-id="3e459-793">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="3e459-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3e459-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3e459-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-796">- ActiveView</span></span><br><span data-ttu-id="3e459-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-797">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-798">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-799">
         - File</span></span><br><span data-ttu-id="3e459-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-800">
         - PdfFile</span></span><br><span data-ttu-id="3e459-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-801">
         - Selection</span></span><br><span data-ttu-id="3e459-802">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-802">
         - Settings</span></span><br><span data-ttu-id="3e459-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-804">Office для iPad</span><span class="sxs-lookup"><span data-stu-id="3e459-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3e459-805">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-806">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-806">- Content</span></span><br><span data-ttu-id="3e459-807">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-810">- ActiveView</span></span><br><span data-ttu-id="3e459-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-811">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-812">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-813">
         - File</span></span><br><span data-ttu-id="3e459-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-814">
         - PdfFile</span></span><br><span data-ttu-id="3e459-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-815">
         - Selection</span></span><br><span data-ttu-id="3e459-816">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-816">
         - Settings</span></span><br><span data-ttu-id="3e459-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-818">Office для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-818">Office apps on Mac</span></span><br><span data-ttu-id="3e459-819">(версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3e459-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3e459-820">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-820">- Content</span></span><br><span data-ttu-id="3e459-821">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-821">
         - TaskPane</span></span><br><span data-ttu-id="3e459-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3e459-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3e459-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3e459-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-826">- ActiveView</span></span><br><span data-ttu-id="3e459-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-827">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-828">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-829">
         - File</span></span><br><span data-ttu-id="3e459-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-830">
         - PdfFile</span></span><br><span data-ttu-id="3e459-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-831">
         - Selection</span></span><br><span data-ttu-id="3e459-832">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-832">
         - Settings</span></span><br><span data-ttu-id="3e459-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-834">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-834">Office 2019 for Mac</span></span><br><span data-ttu-id="3e459-835">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-836">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-836">- Content</span></span><br><span data-ttu-id="3e459-837">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-837">
         - TaskPane</span></span><br><span data-ttu-id="3e459-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-841">- ActiveView</span></span><br><span data-ttu-id="3e459-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-842">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-843">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-844">
         - File</span></span><br><span data-ttu-id="3e459-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-845">
         - PdfFile</span></span><br><span data-ttu-id="3e459-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-846">
         - Selection</span></span><br><span data-ttu-id="3e459-847">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-847">
         - Settings</span></span><br><span data-ttu-id="3e459-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-849">Office 2016 для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-849">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3e459-850">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-851">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-851">- Content</span></span><br><span data-ttu-id="3e459-852">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3e459-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3e459-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3e459-855">- ActiveView</span></span><br><span data-ttu-id="3e459-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3e459-856">
         - CompressedFile</span></span><br><span data-ttu-id="3e459-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-857">
         - DocumentEvents</span></span><br><span data-ttu-id="3e459-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="3e459-858">
         - File</span></span><br><span data-ttu-id="3e459-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3e459-859">
         - PdfFile</span></span><br><span data-ttu-id="3e459-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-860">
         - Selection</span></span><br><span data-ttu-id="3e459-861">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-861">
         - Settings</span></span><br><span data-ttu-id="3e459-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="3e459-863">*&ast; - Добавлены обновления после выпуска.*</span><span class="sxs-lookup"><span data-stu-id="3e459-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="3e459-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="3e459-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3e459-865">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-865">Platform</span></span></th>
    <th><span data-ttu-id="3e459-866">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-866">Extension points</span></span></th>
    <th><span data-ttu-id="3e459-867">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="3e459-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-869">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="3e459-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="3e459-870">- Контент</span><span class="sxs-lookup"><span data-stu-id="3e459-870">- Content</span></span><br><span data-ttu-id="3e459-871">
         - Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-871">
         - TaskPane</span></span><br><span data-ttu-id="3e459-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></span><span class="sxs-lookup"><span data-stu-id="3e459-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3e459-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="3e459-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3e459-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3e459-876">- DocumentEvents</span></span><br><span data-ttu-id="3e459-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="3e459-878">
         - Параметры</span><span class="sxs-lookup"><span data-stu-id="3e459-878">
         - Settings</span></span><br><span data-ttu-id="3e459-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="3e459-880">Project</span><span class="sxs-lookup"><span data-stu-id="3e459-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3e459-881">Платформа</span><span class="sxs-lookup"><span data-stu-id="3e459-881">Platform</span></span></th>
    <th><span data-ttu-id="3e459-882">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3e459-882">Extension points</span></span></th>
    <th><span data-ttu-id="3e459-883">Наборы обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="3e459-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="3e459-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Общие API</b></a></span><span class="sxs-lookup"><span data-stu-id="3e459-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-885">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-885">Office 2019 on Windows</span></span><br><span data-ttu-id="3e459-886">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-887">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-889">- Selection</span></span><br><span data-ttu-id="3e459-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-891">Office 2016 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-891">Office 2016 on Windows</span></span><br><span data-ttu-id="3e459-892">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-893">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-895">- Selection</span></span><br><span data-ttu-id="3e459-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3e459-897">Office 2013 для Windows</span><span class="sxs-lookup"><span data-stu-id="3e459-897">Office 2013 on Windows</span></span><br><span data-ttu-id="3e459-898">(единовременная покупка)</span><span class="sxs-lookup"><span data-stu-id="3e459-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3e459-899">- Область задач</span><span class="sxs-lookup"><span data-stu-id="3e459-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3e459-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3e459-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3e459-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="3e459-901">- Selection</span></span><br><span data-ttu-id="3e459-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3e459-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="3e459-903">См. также</span><span class="sxs-lookup"><span data-stu-id="3e459-903">See also</span></span>

- [<span data-ttu-id="3e459-904">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="3e459-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="3e459-905">Версии Office и наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="3e459-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="3e459-906">Наборы обязательных элементов общего API</span><span class="sxs-lookup"><span data-stu-id="3e459-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="3e459-907">Наборы обязательных элементов для команд надстроек</span><span class="sxs-lookup"><span data-stu-id="3e459-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="3e459-908">Справка по API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="3e459-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="3e459-909">Журнал обновлений для Office 365 профессиональный плюс</span><span class="sxs-lookup"><span data-stu-id="3e459-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="3e459-910">Журнал обновлений Office 2016 и 2019 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="3e459-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="3e459-911">Журнал обновлений Office 2013 ("нажми и работай")</span><span class="sxs-lookup"><span data-stu-id="3e459-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="3e459-912">Журнал обновлений Office 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3e459-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="3e459-913">Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3e459-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="3e459-914">Журнал обновлений Office для Mac</span><span class="sxs-lookup"><span data-stu-id="3e459-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
